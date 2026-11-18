# Confucius: course-grounded tutor architecture

Confucius is a private tutoring agent design for MBA coursework. This public snapshot contains the **actual offline source indexer** used to locate course evidence. The original code was adapted to remove a machine-specific path and archive-specific course notes, keep generated indexes in ignored folders, and reject course-folder escapes. It also includes synthetic tests from the original suite plus a folder-boundary regression. It does not contain course materials, learner records, credentials, conversations, a live agent service, or the scripted tutor shown on my portfolio website.

## Architecture

```mermaid
flowchart LR
  A[Confirmed course] --> B[Course-scoped source manifest]
  B --> C[Offline extraction and source locators]
  C --> D[SQLite FTS5 index]
  D --> E[Search within course]
  E --> F[Read source passage and context]
  F --> G[Tutoring mode and grounded response]
```

The private agent policy defines nine modes: Teach, Quiz, Problem Set Coach, Case Prep, Reading Brief, Draft Review, Rehearse, Plan, and Exam Review. It starts with course confirmation, asks a diagnostic question when understanding is unclear, and uses progressive hints for active graded work. Its source order begins with the current assignment and rubric, then syllabus and instructor material, then course system content. It asks for source locations and labels outside background. These are **designed behaviors**, not evidence that every model response complied or that student learning improved.

`course_index.py` implements the retrieval side. It extracts supported local PDFs, documents, spreadsheets, slides, notebooks, CSV, and HTML without executing embedded code or macros. It records source IDs, hashes, extraction status, gaps, and page or logical locators, then builds a SQLite FTS5 index. `search` uses literal token matching, optionally filtered by an exact course ID. It is keyword search, not embeddings or a mastery model. The tutor workflow requires course selection; the indexer's course filter is optional, so callers must enforce that workflow.

The private system also had a Canvas metadata checker and a local archive. Neither is in this snapshot. The checker could download newly detected files when applied, so this package does not imply that source refreshes are read-only or approval-enforced. The original teaching policy and archived workspace remain private. The tutor skill that informed the policy was credited to Hermes Agent; this package does not claim exclusive authorship of that upstream material.

## Inspect and test

Use Python 3.14 and install the pinned packages in `requirements.lock.txt` in a virtual environment. The tests create disposable synthetic sources, exercise extraction and course separation, and do not need a course account or network connection.

```sh
python -m pip install -r requirements.lock.txt
python -m pytest test_index.py -q
python course_index.py --help
```

To index your own authorized local archive, supply its path explicitly:

```sh
python course_index.py build --archive /path/to/your/archive --output .local-index
python course_index.py search 'example topic' --course 1 --output .local-index
```

The build command expects the original archive manifest format; `test_index.py` constructs a minimal compatible synthetic example. Search opens the generated database read-only. The code does not contact Canvas. Keep actual course files and generated indexes out of Git.

## Evidence limits

The supplied tests cover synthetic extraction and retrieval behavior, including course separation, restricted solution content, and course-folder confinement. They do not evaluate live tutoring, citation accuracy in generated answers, current source freshness, or learning outcomes. The selected indexer code is MIT licensed; see [LICENSE](LICENSE). The related Hermes Agent tutoring skill and private course material are not included.
