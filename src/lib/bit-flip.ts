// .dst files are UTF-8 XML run through a fixed byte-substitution cipher.
// Each .dst byte splits into a high nibble (block) and a low nibble (index):
//
//   xml = BLOCK_BASE[dst >> 4] + PERMUTE[dst & 0xF]
//
// PERMUTE is the same for every block and is self-inverse. The confirmed
// blocks cover XML bytes 31-254, which includes every UTF-8 lead and
// continuation byte, so non-ASCII text round-trips. Block 0x8 carries TAB/LF
// instead; block 0xb (XML 0-30, control characters) is unconfirmed and passes
// through unchanged.
// Reference: https://github.com/Bedz01/dst-codec

const PERMUTE = [13, 12, 15, 14, 9, 8, 11, 10, 5, 4, 7, 6, 1, 0, 3, 2]

// prettier-ignore
const BLOCK_BASE: Record<number, number> = {
  0x0: 127, 0x1: 111, 0x2: 159, 0x3: 143, 0x4: 191, 0x5: 175, 0x6: 223, 0x7: 207,
  0x9: 239, 0xa: 31,  0xc: 63,  0xd: 47,  0xe: 95,  0xf: 79,
}

// Block 0x8: AutoCAD writes LF as 131, ARES as 135; both write TAB as 134.
const WHITESPACE: Record<number, number> = { 131: 10, 134: 9, 135: 10 }

const decodeMap = new Uint8Array(256).map((_, i) => i)
const encodeMap = new Uint8Array(256).map((_, i) => i)

for (let dst = 0; dst < 256; dst++) {
  const base = BLOCK_BASE[dst >> 4]
  if (base === undefined) continue
  const xml = base + PERMUTE[dst & 0xf]
  decodeMap[dst] = xml
  encodeMap[xml] = dst
}
for (const [dst, xml] of Object.entries(WHITESPACE)) decodeMap[Number(dst)] = xml
encodeMap[9] = 134
encodeMap[10] = 131

function flipBits(byteValue: number, reverse = false): number {
  return (reverse ? encodeMap : decodeMap)[byteValue]
}

/**
 * Decodes a DST file by flipping bits to reveal XML content
 */
export function decodeDST(arrayBuffer: ArrayBuffer): string {
  const view = new DataView(arrayBuffer)
  for (let i = 0; i < view.byteLength; i++) {
    const byteValue = view.getUint8(i)
    const flippedValue = flipBits(byteValue)
    view.setUint8(i, flippedValue)
  }
  const textDecoder = new TextDecoder("utf-8")
  return textDecoder.decode(view.buffer)
}

/**
 * Encodes XML content into DST format by flipping bits
 *
 * Line endings are normalised to LF first: CR has no .dst byte, and passed
 * through raw it decodes as DEL, which breaks the file. Firefox's
 * XMLSerializer on Windows writes CRLF after the XML declaration. XML parsers
 * normalise line endings to LF anyway, so nothing is lost.
 */
export function encodeDST(xmlString: string): ArrayBuffer {
  const textEncoder = new TextEncoder()
  const normalised = xmlString.replace(/\r\n?/g, "\n")
  const view = new DataView(textEncoder.encode(normalised).buffer)
  for (let i = 0; i < view.byteLength; i++) {
    const byteValue = view.getUint8(i)
    const flippedValue = flipBits(byteValue, true)
    view.setUint8(i, flippedValue)
  }
  return view.buffer
}
