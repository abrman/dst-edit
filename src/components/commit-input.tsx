import { useEffect, useState } from "react";
import type React from "react";
import { cn } from "@/lib/utils";

interface CommitInputProps extends Omit<React.ComponentProps<"input">, "value" | "onChange"> {
  value: string;
  /** Called once when an edit is finished (blur or Enter), so each edit is a single undo step. */
  onCommit: (value: string) => void;
  /** Cleans the typed value before it's committed (e.g. replacing characters AutoCAD doesn't allow). */
  normalize?: (value: string) => string;
}

/** A text input that keeps its own draft and commits on blur or Enter; Escape cancels. */
export function CommitInput({ value, onCommit, normalize, onKeyDown, onBlur, className, ...props }: CommitInputProps) {
  const [draft, setDraft] = useState(value);
  useEffect(() => setDraft(value), [value]);

  const commit = () => {
    const next = normalize ? normalize(draft) : draft;
    setDraft(next);
    if (next !== value) onCommit(next);
  };

  return (
    <input
      {...props}
      value={draft}
      onChange={(e) => setDraft(e.target.value)}
      onBlur={(e) => {
        commit();
        onBlur?.(e);
      }}
      onKeyDown={(e) => {
        // Blurring commits, so Enter and the blur that follows don't commit twice
        if (e.key === "Enter") e.currentTarget.blur();
        if (e.key === "Escape") {
          setDraft(value);
          // Blur after the reset renders, so blurring doesn't commit the discarded draft
          const input = e.currentTarget;
          requestAnimationFrame(() => input.blur());
        }
        onKeyDown?.(e);
      }}
      className={cn(
        "h-8 w-full min-w-0 rounded-sm border border-transparent bg-transparent px-2 text-sm outline-none transition-colors",
        "placeholder:text-muted-foreground/60 hover:border-input focus:border-ring focus:bg-background focus:ring-2 focus:ring-ring/30",
        className,
      )}
    />
  );
}
