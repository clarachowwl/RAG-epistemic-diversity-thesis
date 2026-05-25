# Retrieval Depth and Epistemic Diversity in Cultural Heritage RAG Systems

**Master's Thesis — University of Amsterdam, 2026**

> *"Retrieval Depth and Epistemic Diversity in Cultural Heritage RAG Systems"*

---

## Overview

This repository contains all code used in the thesis examining how retrieval depth configurations in RAG (Retrieval-Augmented Generation) systems affect epistemic diversity in Dutch colonial cultural heritage search.

The study tests three retrieval depth conditions (k=3, k=10, k=20) across a three-layer corpus of 242 documents, using 20 researcher-designed queries and 15 user-generated queries.

---

## Repository Structure

```
rag-epistemic-diversity-thesis/
├── README.md
├── requirements.txt
├── corpus/
│   ├── layer1_rijksmuseum.py       # Layer 1: Rijksmuseum API
│   ├── layer2_europeana.py         # Layer 2: Europeana API
│   └── layer3_wikipedia.py         # Layer 3: Wikipedia critical sources
├── pipeline/
│   └── rag_pipeline.py             # Main RAG experiment pipeline
└── annotation/
    └── ai_annotator.py             # AI-as-judge annotation script
```

---

## Research Design

| Component | Details |
|---|---|
| Corpus | 242 documents (Layer 1: 74, Layer 2: 145, Layer 3: 23) |
| Researcher queries | 20 |
| User-generated queries | 15 (from 6 participants) |
| k-values tested | 3, 10, 20 |
| Embedding model | text-embedding-ada-002 |
| Generation model | GPT-4o |
| Annotation model | GPT-4o (AI-as-judge) |

---

## Three-layer Corpus

| Layer | Source | Documents | Epistemic Position |
|---|---|---|---|
| Layer 1 | Rijksmuseum API | 74 | Institutional |
| Layer 2 | Europeana API | 145 | Cross-institutional |
| Layer 3 | Wikipedia | 23 | Critical/decolonial |

---

## Installation

```bash
pip install -r requirements.txt
```

---

## Usage

### Step 1: Build the corpus

```bash
# Build Layer 1 (Rijksmuseum)
python corpus/layer1_rijksmuseum.py

# Build Layer 2 (Europeana)
python corpus/layer2_europeana.py

# Build Layer 3 (Wikipedia)
python corpus/layer3_wikipedia.py
```

### Step 2: Run RAG pipeline

```bash
# Add your OpenAI API key to rag_pipeline.py first
python pipeline/rag_pipeline.py
```

### Step 3: Run AI annotation

```bash
# Add your OpenAI API key to ai_annotator.py first
python annotation/ai_annotator.py
```

---

## Configuration

You need to add your own OpenAI API key before running the pipeline or annotation scripts.

```python
# In rag_pipeline.py and ai_annotator.py
OPENAI_API_KEY = "your-openai-api-key-here"
```

---

## Key Findings

1. **RQ1:** Layer 3 document count increased by 361% between k=3 and k=20 for researcher-designed queries
2. **RQ2:** Institutional epistemic stance (ES-1) decreased from 12/20 to 2/20 between k=3 and k=20
3. **RQ3:** Epistemic filtering operates across three pipeline stages: query formulation, retrieval, and generation

---

## Citation

If you use this code, please cite:

```
[Author]. (2026). Retrieval Depth and Epistemic Diversity in 
Cultural Heritage RAG Systems. Master's thesis, 
University of Amsterdam.
```

---

## License

MIT License — see LICENSE file for details.
