import os
from pathlib import Path

book_dir = Path('book1')
files = ['prologue.md'] + [f'chapter_{i:02d}.md' for i in range(1, 15)] + ['epilogue.md']

total_words = 0
for f in files:
    file_path = book_dir / f
    if file_path.exists():
        content = file_path.read_text(encoding='utf-8')
        # Remove markdown headers and count words
        words = len(content.split())
        total_words += words
        print(f"{f:20s}: {words:5d} words")
    else:
        print(f"{f:20s}: NOT FOUND")

print(f"\n{'Total':20s}: {total_words:5d} words")
print(f"\nAverage per chapter: {total_words // len([f for f in files if (book_dir / f).exists()]):.0f} words")

