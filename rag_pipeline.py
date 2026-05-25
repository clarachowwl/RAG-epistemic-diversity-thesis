"""
RAG Pipeline — Epistemic Diversity Experiment
==============================================
Varies retrieval depth (k=3, k=10, k=20) and measures
epistemic diversity in generated summaries.

Usage:
    pip install openai pandas openpyxl numpy scikit-learn
    python rag_pipeline.py

Note: Add your OpenAI API key before running.
"""

import json
import time
import numpy as np
import pandas as pd
import openpyxl
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from openpyxl.utils import get_column_letter
from datetime import date
from sklearn.metrics.pairwise import cosine_similarity
from openai import OpenAI

# ─────────────────────────────────────────
# CONFIGURATION
# ─────────────────────────────────────────
OPENAI_API_KEY   = "your-openai-api-key-here"   # <-- Add your key
CORPUS_FILE      = "unified_corpus.xlsx"
OUTPUT_FILE      = "rag_results.xlsx"
EMBEDDING_MODEL  = "text-embedding-ada-002"
GENERATION_MODEL = "gpt-4o"
K_VALUES         = [3, 10, 20]

QUERIES = [
    # Original 8
    "What does Dutch colonial art tell us about the relationship between the Netherlands and Indonesia?",
    "Whose perspectives are missing from European museum collections about colonial history?",
    "How did Dutch colonialism shape the cultural identity of Suriname?",
    "How do colonial portraits reflect power dynamics between colonisers and colonised?",
    "What are the experiences of enslaved people in Dutch colonial history, and how are they documented?",
    "How did the VOC exploit Indonesian and Surinamese communities, and what traces remain in cultural collections?",
    "How is Batavia remembered differently by Dutch colonisers and Indonesian communities?",
    "How have Indonesian communities resisted or responded to Dutch colonial cultural impositions?",
    # Repatriation
    "What arguments exist for and against the repatriation of colonial objects from Dutch museums to Indonesia and Suriname?",
    "How do Dutch museums justify their continued ownership of objects acquired during the colonial period?",
    "What do source communities in Indonesia and Suriname say about cultural objects held in Dutch collections?",
    # Gender / Marginalised voices
    "How are women and gender represented in Dutch colonial heritage collections?",
    "What is the role of enslaved and colonised women in Dutch colonial history, and how is this documented?",
    "How do Dutch colonial collections represent indigenous knowledge systems and practices?",
    # Memory / Commemoration
    "How is Dutch colonial violence remembered and commemorated in the Netherlands and in Indonesia?",
    "What is the significance of Dutch colonial exhibitions in shaping public understanding of colonial history?",
    "How do Dutch cultural institutions acknowledge or avoid addressing the legacy of slavery?",
    # Decolonisation
    "What steps have Dutch museums taken to decolonise their collections and narratives?",
    "How do postcolonial scholars critique the way Dutch museums present colonial history?",
    "What would a decolonised Dutch cultural heritage collection look like, according to affected communities?",
]

client = OpenAI(api_key=OPENAI_API_KEY)

# ─────────────────────────────────────────
# STEP 1: LOAD CORPUS
# ─────────────────────────────────────────
def load_corpus(filepath):
    df = pd.read_excel(filepath, sheet_name=0)
    df.columns = [c.lower().replace(" ", "_") for c in df.columns]
    df = df[df["description"].astype(str).str.len() > 20].reset_index(drop=True)
    print(f"Loaded corpus: {len(df)} documents")
    print(f"  Layer 1: {len(df[df.epistemic_layer == 1])}")
    print(f"  Layer 2: {len(df[df.epistemic_layer == 2])}")
    print(f"  Layer 3: {len(df[df.epistemic_layer == 3])}")
    return df

# ─────────────────────────────────────────
# STEP 2: EMBED CORPUS
# ─────────────────────────────────────────
def embed_texts(texts, batch_size=50):
    all_embeddings = []
    for i in range(0, len(texts), batch_size):
        batch = texts[i:i+batch_size]
        print(f"  Embedding batch {i//batch_size + 1}/{(len(texts)-1)//batch_size + 1}...")
        response = client.embeddings.create(model=EMBEDDING_MODEL, input=batch)
        all_embeddings.extend([item.embedding for item in response.data])
        time.sleep(0.5)
    return np.array(all_embeddings)


def build_vector_store(df):
    print("\nBuilding vector store...")
    embeddings = embed_texts(df["description"].astype(str).tolist())
    print(f"Embedded {len(embeddings)} documents")
    return embeddings

# ─────────────────────────────────────────
# STEP 3: RETRIEVE TOP-K
# ─────────────────────────────────────────
def retrieve(query, corpus_embeddings, df, k):
    response       = client.embeddings.create(model=EMBEDDING_MODEL, input=[query])
    query_emb      = np.array(response.data[0].embedding).reshape(1, -1)
    similarities   = cosine_similarity(query_emb, corpus_embeddings)[0]
    top_k_indices  = np.argsort(similarities)[::-1][:k]
    retrieved      = df.iloc[top_k_indices].copy()
    retrieved["similarity_score"] = similarities[top_k_indices]
    return retrieved

# ─────────────────────────────────────────
# STEP 4: GENERATE SUMMARY
# ─────────────────────────────────────────
def generate_summary(query, retrieved_docs):
    context_parts = []
    for _, doc in retrieved_docs.iterrows():
        source = doc.get("source_institution", "Unknown source")
        title  = doc.get("title", "Untitled")
        text   = doc.get("description", "")
        context_parts.append(f"[Source: {source} | {title}]\n{text}")

    context  = "\n\n---\n\n".join(context_parts)
    prompt   = f"""You are a cultural heritage research assistant.
Using ONLY the provided sources, write a comprehensive summary that answers the question below.
Include perspectives from all sources provided. Note any diversity or tension between perspectives.

Question: {query}

Sources:
{context}

Write a 200-300 word summary based on these sources."""

    response = client.chat.completions.create(
        model=GENERATION_MODEL,
        messages=[{"role": "user", "content": prompt}],
        temperature=0.3,
        max_tokens=500,
    )
    return response.choices[0].message.content.strip()

# ─────────────────────────────────────────
# STEP 5: MEASURE DIVERSITY
# ─────────────────────────────────────────
def measure_diversity(retrieved_docs, corpus_embeddings, df):
    layer_counts  = retrieved_docs["epistemic_layer"].value_counts().to_dict()
    institutions  = retrieved_docs["source_institution"].dropna().unique()
    stances       = retrieved_docs["epistemic_stance"].dropna().unique()
    topics        = retrieved_docs["topic"].dropna().unique()

    indices = retrieved_docs.index.tolist()
    if len(indices) > 1:
        retrieved_emb = corpus_embeddings[indices]
        centroid      = retrieved_emb.mean(axis=0)
        variance      = float(np.mean(np.linalg.norm(retrieved_emb - centroid, axis=1)))
    else:
        variance = 0.0

    return {
        "layer_1_count":    layer_counts.get(1, 0),
        "layer_2_count":    layer_counts.get(2, 0),
        "layer_3_count":    layer_counts.get(3, 0),
        "n_institutions":   len(institutions),
        "n_stances":        len(stances),
        "n_topics":         len(topics),
        "embedding_variance": round(variance, 4),
        "layer3_present":   "YES" if layer_counts.get(3, 0) > 0 else "NO",
        "institutions_list": "; ".join(institutions[:5]),
        "stances_list":     "; ".join(stances),
    }

# ─────────────────────────────────────────
# STEP 6: RUN EXPERIMENT
# ─────────────────────────────────────────
def run_experiment(df, corpus_embeddings):
    results = []
    for query in QUERIES:
        print(f"\nQuery: {query[:60]}...")
        for k in K_VALUES:
            print(f"  k={k}: retrieving...")
            retrieved = retrieve(query, corpus_embeddings, df, k)
            diversity = measure_diversity(retrieved, corpus_embeddings, df)
            print(f"  k={k}: generating summary...")
            summary   = generate_summary(query, retrieved)
            results.append({
                "query":   query,
                "k_value": k,
                "summary": summary,
                "retrieved_titles": "; ".join(retrieved["title"].astype(str).tolist()[:5]),
                **diversity,
            })
            print(f"  k={k}: done (L3={diversity['layer3_present']}, var={diversity['embedding_variance']})")
            time.sleep(1)
    return results

# ─────────────────────────────────────────
# STEP 7: SAVE RESULTS
# ─────────────────────────────────────────
RESULT_COLUMNS = [
    ("query",               "Query",                45),
    ("k_value",             "k Value",               8),
    ("layer_1_count",       "Layer 1 Count",        12),
    ("layer_2_count",       "Layer 2 Count",        12),
    ("layer_3_count",       "Layer 3 Count",        12),
    ("layer3_present",      "Layer 3 Present?",     14),
    ("n_institutions",      "# Institutions",       14),
    ("n_stances",           "# Stances",            10),
    ("n_topics",            "# Topics",             10),
    ("embedding_variance",  "Embedding Variance",   16),
    ("institutions_list",   "Institutions",         35),
    ("stances_list",        "Stances",              28),
    ("retrieved_titles",    "Retrieved Titles",     50),
    ("summary",             "Generated Summary",    80),
]
K_COLOURS = {3: "FCE4D6", 10: "DDEBF7", 20: "E2EFDA"}


def save_results(results, filename):
    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = "RAG Results"
    ws.freeze_panes = "A2"
    thin   = Side(style="thin", color="AAAAAA")
    BORDER = Border(left=thin, right=thin, top=thin, bottom=thin)

    for ci, (field, label, width) in enumerate(RESULT_COLUMNS, 1):
        cell = ws.cell(1, ci, label)
        cell.font      = Font(bold=True, color="FFFFFF", size=10, name="Arial")
        cell.fill      = PatternFill("solid", fgColor="1F3864")
        cell.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)
        cell.border    = BORDER
        ws.column_dimensions[get_column_letter(ci)].width = width
    ws.row_dimensions[1].height = 22

    for ri, result in enumerate(results, 2):
        k        = result.get("k_value", 3)
        row_fill = PatternFill("solid", fgColor=K_COLOURS.get(k, "FFFFFF"))
        for ci, (field, label, _) in enumerate(RESULT_COLUMNS, 1):
            cell = ws.cell(ri, ci, result.get(field, ""))
            cell.font      = Font(size=10, name="Arial")
            cell.fill      = row_fill
            cell.alignment = Alignment(vertical="top", wrap_text=(field == "summary"))
            cell.border    = BORDER
        ws.row_dimensions[ri].height = 80 if result.get("summary") else 28

    wb.save(filename)
    print(f"\nSaved: {filename} ({len(results)} rows)")


if __name__ == "__main__":
    if OPENAI_API_KEY == "your-openai-api-key-here":
        print("Please add your OpenAI API key first.")
    else:
        print("Starting RAG Epistemic Diversity Experiment\n")
        df               = load_corpus(CORPUS_FILE)
        corpus_embeddings = build_vector_store(df)
        results          = run_experiment(df, corpus_embeddings)
        save_results(results, OUTPUT_FILE)
        print("\nExperiment complete!")
