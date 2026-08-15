"""A long module, for timing the renderer and for checking gutter alignment.

Nothing here is subtle. The point is volume: enough lines that the line-number
gutter has to stay aligned past 100, and enough coloured spans that the
per-character COM round trips in the renderer are measurable.
"""

from __future__ import annotations

import json
import re
from collections import defaultdict

WORD_RE = re.compile(r"[A-Za-z']+")
STOPWORDS = {
    "a", "an", "and", "as", "at", "but", "by", "for", "from", "in", "is",
    "it", "of", "on", "or", "the", "to", "was", "were", "with",
}
DEFAULT_ENCODING = "utf-8"
MAX_REPORT_ROWS = 25


class TokenStream:
    """Splits text into lowercase words, skipping stopwords."""

    def __init__(self, text, stopwords=None):
        self.text = text
        self.stopwords = stopwords if stopwords is not None else STOPWORDS
        self._count = 0

    def __iter__(self):
        for match in WORD_RE.finditer(self.text):
            word = match.group(0).lower()
            if word in self.stopwords:
                continue
            if len(word) < 2:
                continue
            self._count += 1
            yield word

    def __len__(self):
        return self._count

    def __repr__(self):
        return f"TokenStream({len(self.text)} chars, {self._count} tokens)"


class Frequencies:
    """Counts, ranks and reports word frequencies."""

    def __init__(self):
        self.counts = defaultdict(int)
        self.total = 0

    def add(self, word, weight=1):
        if weight <= 0:
            raise ValueError(f"weight must be positive, got {weight}")
        self.counts[word] += weight
        self.total += weight
        return self

    def extend(self, words):
        for word in words:
            self.add(word)
        return self

    def most_common(self, limit=MAX_REPORT_ROWS):
        ranked = sorted(self.counts.items(), key=lambda kv: (-kv[1], kv[0]))
        return ranked[:limit]

    def share(self, word):
        if self.total == 0:
            return 0.0
        return self.counts[word] / self.total

    def to_json(self, indent=2):
        payload = {
            "total": self.total,
            "unique": len(self.counts),
            "top": [{"word": w, "count": c} for w, c in self.most_common()],
        }
        return json.dumps(payload, indent=indent, sort_keys=True)

    def __contains__(self, word):
        return word in self.counts

    def __len__(self):
        return len(self.counts)


def read_text(path, encoding=DEFAULT_ENCODING):
    """Read a file, returning an empty string rather than raising."""
    try:
        with open(path, encoding=encoding) as handle:
            return handle.read()
    except FileNotFoundError:
        return ""
    except UnicodeDecodeError:
        with open(path, encoding="latin-1") as handle:
            return handle.read()


def analyse(text, stopwords=None):
    stream = TokenStream(text, stopwords)
    freq = Frequencies()
    freq.extend(stream)
    return freq


def format_row(word, count, total, width=20):
    label = word.ljust(width)[:width]
    pct = 100.0 * count / total if total else 0.0
    bar = "#" * int(pct)
    return f"{label} {count:>6} {pct:>6.2f}%  {bar}"


def report(freq, limit=MAX_REPORT_ROWS):
    lines = []
    rows = freq.most_common(limit)
    if not rows:
        return "no words counted"
    lines.append(f"{len(freq)} unique words, {freq.total} total")
    lines.append("-" * 52)
    for word, count in rows:
        lines.append(format_row(word, count, freq.total))
    lines.append("-" * 52)
    return "\n".join(lines)


def compare(left, right, limit=10):
    """Words that are much more common in left than in right."""
    deltas = {}
    for word in left.counts:
        delta = left.share(word) - right.share(word)
        if delta > 0:
            deltas[word] = delta
    ranked = sorted(deltas.items(), key=lambda kv: -kv[1])
    return ranked[:limit]


def main(paths):
    combined = Frequencies()
    per_file = {}
    for path in paths:
        text = read_text(path)
        if not text:
            print(f"skipping empty or missing file: {path}")
            continue
        freq = analyse(text)
        per_file[path] = freq
        combined.extend(freq.counts.keys())
        print(report(freq))
    if len(per_file) == 2:
        first, second = per_file.values()
        for word, delta in compare(first, second):
            print(f"{word:<20} {delta:+.4f}")
    return combined


if __name__ == "__main__":
    import sys

    result = main(sys.argv[1:])
    print(result.to_json())
