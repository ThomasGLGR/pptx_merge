# pptx_merge

Merges two PowerPoint (`.pptx`) files into one, **100% natively on Linux**, without relying on `win32com` or Microsoft Office, and completely free of proprietary dependencies.

The script directly manipulates the **OOXML** format (a `.pptx` file is essentially a ZIP archive of XML parts connected by `.rels` relationship files). It uses Presentation A as the base and imports all slides from B along with their entire dependency chain, cleanly renaming each part to prevent collisions.

---

## Features

* **100% Faithful Merging**: Each slide set retains its own master, layouts, and theme. As a result, the appearance of the slides from B remains identical to how they looked in B.
* **Media Deduplication**: If an identical image exists in both A and B (compared using SHA-256 hashing), it is stored only once.
* **Automatic Scaling**: If A and B use different slide sizes, the imported content is automatically resized (both geometry and typography) to perfectly fit the target canvas.
* **Slide Selection**: Choose to keep only specific slides from A and/or B, and control exactly where B's slides are inserted.
* **Notes Management**: Speaker notes from B are linked to A's notes master if it exists (otherwise, B's notes master is imported).

---

## Prerequisites

* Python **3.8+**
* [`lxml`](https://www.google.com/search?q=%5Bhttps://lxml.de/%5D(https://lxml.de/))

`python-pptx` is **not** required by the script; it is only used in tests to generate and verify files.

---

## Installation

```bash
pip install lxml

```

Then just grab `pptx_merge.py`; no other files are needed.

---

## Command Line Usage

Simple merge (slides from A, followed by slides from B):

```bash
python pptx_merge.py A.pptx B.pptx -o fusion.pptx

```

Choose the final slide size when A and B differ:

```bash
python pptx_merge.py A.pptx B.pptx -o fusion.pptx --target-size smallest

```

Keep only specific slides and insert them at a precise position:

```bash
# Imports slides 1, 2, and 5-7 from B, inserted at the very beginning of A
python pptx_merge.py A.pptx B.pptx -o fusion.pptx --slides-b 1,2,5-7 --insert-index 0

```

### Options

| Option | Values | Default | Description |
| --- | --- | --- | --- |
| `-o`, `--output` | path to `.pptx` | `fusion.pptx` | Output file. |
| `--target-size` | `first`, `second`, `largest`, `smallest` | `first` | Final slide size if A and B differ. `first` keeps A intact and scales B up/down. |
| `--slides-a` | e.g., `1,3,5-8` | all | Slides from A to keep (1-indexed). |
| `--slides-b` | e.g., `2,4` | all | Slides from B to import. |
| `--insert-index` | integer | end of A | Insertion position for B's slides within A's slide list (`0` = at the beginning). |

---

## Usage as a Library

```python
from pptx_merge import merge_pptx, merge_specific_slides

# Simple merge
merge_pptx("A.pptx", "B.pptx", "fusion.pptx", target_size="first")

# Selective merge: keep slides 1 and 3 from A, import slides 2, 4, and 5
# from B, and insert them right after the first retained slide of A.
merge_specific_slides(
    file_a="A.pptx",
    file_b="B.pptx",
    slides_a=[1, 3],
    slides_b=[2, 4, 5],
    insert_index=1,
    output="fusion_specifique.pptx",
    target_size="first",
)

```

The `target_size` parameter accepts:

* `"first"` — size of A (default); slides from B are scaled.
* `"second"` — size of B; slides from A are scaled.
* `"largest"` / `"smallest"` — the larger / smaller of the two (by surface area).
