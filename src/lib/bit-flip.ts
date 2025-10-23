// prettier-ignore
const flipMap: Record<number, number> = {
  131: 10,
  134: 9,
  172: 32,
  175: 33,
  174: 34,
  169: 35,
  168: 36,
  171: 37,
  170: 38,
  165: 39,
  164: 40,
  167: 41,
  166: 42,
  161: 43,
  160: 44,
  163: 45,
  162: 46,
  221: 47,
  220: 48,
  223: 49,
  222: 50,
  217: 51,
  216: 52,
  219: 53,
  218: 54,
  213: 55,
  212: 56,
  215: 57,
  214: 58,
  209: 59,
  208: 60,
  211: 61,
  210: 62,
  205: 63,
  204: 64,
  207: 65,
  206: 66,
  201: 67,
  200: 68,
  203: 69,
  202: 70,
  197: 71,
  196: 72,
  199: 73,
  198: 74,
  193: 75,
  192: 76,
  195: 77,
  194: 78,
  253: 79,
  252: 80,
  255: 81,
  254: 82,
  249: 83,
  248: 84,
  251: 85,
  250: 86,
  245: 87,
  244: 88,
  247: 89,
  246: 90,
  241: 91,
  240: 92,
  243: 93,
  242: 94,
  237: 95,
  236: 96,
  239: 97,
  238: 98,
  233: 99,
  232: 100,
  235: 101,
  234: 102,
  229: 103,
  228: 104,
  231: 105,
  230: 106,
  225: 107,
  224: 108,
  227: 109,
  226: 110,
  29: 111,
  28: 112,
  31: 113,
  30: 114,
  25: 115,
  24: 116,
  27: 117,
  26: 118,
  21: 119,
  20: 120,
  23: 121,
  22: 122,
}

const reversedMap = Object.keys(flipMap).reduce(
  (map, key) => ({ ...map, [flipMap[Number(key)]]: Number(key) }),
  {} as Record<number, number>,
)

function flipBits(byteValue: number, reverse = false): number {
  const map = reverse ? reversedMap : flipMap
  return map[byteValue] !== undefined ? map[byteValue] : byteValue
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
 */
export function encodeDST(xmlString: string): ArrayBuffer {
  const textEncoder = new TextEncoder()
  const view = new DataView(textEncoder.encode(xmlString).buffer)
  for (let i = 0; i < view.byteLength; i++) {
    const byteValue = view.getUint8(i)
    const flippedValue = flipBits(byteValue, true)
    view.setUint8(i, flippedValue)
  }
  return view.buffer
}
