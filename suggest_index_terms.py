"""Suggest per-document index terms and preserve review decisions.

Run from any directory with:
    python3 suggest_index_terms.py
"""

import argparse
import csv
import re
from collections import Counter
from pathlib import Path

from docx import Document
from nltk import pos_tag_sents, sent_tokenize, word_tokenize
from nltk.chunk import ne_chunk_sents
from nltk.corpus import brown
from nltk.tree import Tree
from sklearn.feature_extraction.text import TfidfVectorizer

ENTITY_LABELS = {"FACILITY", "GPE", "GSP", "LOCATION", "ORGANIZATION", "PERSON"}


def natural_sort_key(path: Path) -> tuple[tuple[int, int | str], ...]:
    parts = re.split(r"(\d+)", path.name)
    return tuple(
        (0, int(part)) if part.isdigit() else (1, part.casefold())
        for part in parts
    )


def extract_text(path: Path) -> str:
    document = Document(path)
    parts = [paragraph.text for paragraph in document.paragraphs]
    parts.extend(
        cell.text
        for table in document.tables
        for row in table.rows
        for cell in row.cells
    )
    return "\n".join(part for part in parts if part.strip())


def extract_named_entities(text: str) -> list[tuple[str, str]]:
    entities = []

    def add_entity(tokens: list[str], entity_type: str | None) -> None:
        if entity_type is None:
            return
        entity = " ".join(tokens)
        entity = re.sub(r"\s+([,.;:!?])", r"\1", entity).strip()
        if entity:
            entities.append((entity, entity_type))

    try:
        tokenized_sentences = [
            word_tokenize(sentence) for sentence in sent_tokenize(text)
        ]
        tagged_sentences = pos_tag_sents(tokenized_sentences)

        for chunked in ne_chunk_sents(tagged_sentences):
            entity_tokens = []
            entity_type = None

            for item in chunked:
                if isinstance(item, Tree) and item.label() in ENTITY_LABELS:
                    if entity_type is not None and item.label() != entity_type:
                        add_entity(entity_tokens, entity_type)
                        entity_tokens = []
                    entity_type = item.label()
                    entity_tokens.extend(token for token, _ in item.leaves())
                else:
                    add_entity(entity_tokens, entity_type)
                    entity_tokens = []
                    entity_type = None

            add_entity(entity_tokens, entity_type)
    except LookupError as error:
        raise RuntimeError(
            "NLTK's English NER data is missing. Install it with:\n"
            "python3 -m nltk.downloader punkt_tab "
            "averaged_perceptron_tagger_eng maxent_ne_chunker_tab words"
        ) from error
    return entities


def candidate_key(document: str, term: str) -> tuple[str, str]:
    normalized_term = " ".join(term.split()).casefold()
    return document.strip().casefold(), normalized_term


def load_review_values(path: Path) -> dict[tuple[str, str], dict[str, object]]:
    if not path.is_file():
        return {}

    with path.open("r", newline="", encoding="utf-8-sig") as csvfile:
        reader = csv.DictReader(csvfile)
        fieldnames = {
            name.strip().casefold(): name
            for name in (reader.fieldnames or [])
            if name is not None
        }
        if "keep" not in fieldnames and "source_candidate_term" not in fieldnames:
            return {}
        if "document" not in fieldnames or "candidate_term" not in fieldnames:
            raise ValueError(
                f"{path} has review data but is missing document or candidate_term."
            )

        review_values = {}
        for line_number, row in enumerate(reader, start=2):
            document = (row.get(fieldnames["document"]) or "").strip()
            term = (row.get(fieldnames["candidate_term"]) or "").strip()
            if not document or not term:
                raise ValueError(
                    f"{path} has a blank document or candidate_term on row {line_number}."
                )

            source_column = fieldnames.get("source_candidate_term")
            source_term = (row.get(source_column) or "").strip() if source_column else ""
            key = candidate_key(document, source_term or term)

            review_value: dict[str, object] = {"candidate_term": term}
            if "keep" in fieldnames:
                raw_value = (row.get(fieldnames["keep"]) or "").strip().casefold()
                if raw_value in {"", "false", "0", "no"}:
                    review_value["Keep"] = False
                elif raw_value in {"true", "1", "yes"}:
                    review_value["Keep"] = True
                else:
                    raise ValueError(
                        f"{path} has an invalid Keep value on row {line_number}: "
                        f"{raw_value!r}"
                    )

            if key in review_values and review_values[key] != review_value:
                raise ValueError(
                    f"{path} has conflicting review values for {document!r}, "
                    f"{source_term or term!r}."
                )
            review_values[key] = review_value

    return review_values


def main() -> None:
    parser = argparse.ArgumentParser(
        description=(
            "Create Brown-background TF-IDF, trigram, and named-entity candidates "
            "for each DOCX."
        )
    )
    parser.add_argument(
        "--directory",
        type=Path,
        default=Path(__file__).resolve().parent,
        help="Folder containing the DOCX files (defaults to this script's folder).",
    )
    parser.add_argument(
        "--output",
        type=Path,
        help=(
            "CSV output path (defaults to index_candidates.csv in the input folder); "
            "existing Keep values in this file are preserved."
        ),
    )
    parser.add_argument(
        "--top-n",
        type=int,
        default=30,
        help=(
            "Maximum number of TF-IDF candidates per document (default: 30)."
        ),
    )
    parser.add_argument(
        "--entity-top-n",
        type=int,
        default=20,
        help=(
            "Maximum number of additional named-entity candidates per document "
            "(default: 20)."
        ),
    )
    parser.add_argument(
        "--trigram-top-n",
        type=int,
        default=10,
        help=(
            "Maximum number of additional three-word TF-IDF candidates per "
            "document (default: 10)."
        ),
    )
    args = parser.parse_args()

    if args.top_n < 1:
        parser.error("--top-n must be at least 1")
    if args.entity_top_n < 0:
        parser.error("--entity-top-n cannot be negative")
    if args.trigram_top_n < 0:
        parser.error("--trigram-top-n cannot be negative")

    directory = args.directory.expanduser().resolve()
    if not directory.is_dir():
        parser.error(f"not a directory: {directory}")

    documents = sorted(
        (
            path
            for path in directory.glob("*.docx")
            if not path.name.startswith("~$")
        ),
        key=natural_sort_key,
    )
    if not documents:
        parser.error(f"no .docx files found in {directory}")

    texts = [extract_text(path) for path in documents]
    if not any(text.strip() for text in texts):
        parser.error("the DOCX files contain no extractable body text")

    try:
        background_texts = [brown.raw(fileid) for fileid in brown.fileids()]
    except LookupError as error:
        raise RuntimeError(
            "The NLTK Brown corpus is missing. Install it with:\n"
            "python3 -m nltk.downloader brown"
        ) from error

    vectorizer = TfidfVectorizer(
        stop_words="english",
        ngram_range=(1, 3),
    )
    vectorizer.fit([*background_texts, *texts])
    tfidf = vectorizer.transform(texts)
    terms = vectorizer.get_feature_names_out()
    analyze = vectorizer.build_analyzer()
    named_entities = [extract_named_entities(text) for text in texts]

    output = args.output or directory / "index_candidates.csv"
    output = output.expanduser().resolve()
    review_values = load_review_values(output)
    output.parent.mkdir(parents=True, exist_ok=True)

    with output.open("w", newline="", encoding="utf-8-sig") as csvfile:
        writer = csv.writer(csvfile)
        writer.writerow(
            [
                "document",
                "candidate_term",
                "tfidf_score",
                "occurrences",
                "rank",
                "source",
                "entity_type",
                "Keep",
                "source_candidate_term",
            ]
        )

        for row_index, (path, text) in enumerate(zip(documents, texts)):
            scores = tfidf[row_index].toarray().ravel()
            candidate_indices = sorted(
                (index for index, score in enumerate(scores) if score > 0),
                key=lambda index: (-scores[index], terms[index]),
            )[: args.top_n]
            counts = Counter(analyze(text))
            all_scores = {
                terms[index].casefold(): scores[index]
                for index, score in enumerate(scores)
                if score > 0
            }
            candidates = []

            for rank, index in enumerate(candidate_indices, start=1):
                term = terms[index]
                candidates.append(
                    {
                        "term": term,
                        "score": scores[index],
                        "occurrences": counts[term],
                        "source": "TF-IDF",
                        "entity_type": "",
                    }
                )

            tfidf_candidates = {
                candidate["term"].casefold(): candidate for candidate in candidates
            }
            trigram_indices = sorted(
                (
                    index
                    for index, score in enumerate(scores)
                    if score > 0 and len(terms[index].split()) == 3
                ),
                key=lambda index: (-scores[index], terms[index]),
            )[: args.trigram_top_n]
            for index in trigram_indices:
                term = terms[index]
                if term.casefold() not in tfidf_candidates:
                    candidate = {
                        "term": term,
                        "score": scores[index],
                        "occurrences": counts[term],
                        "source": "TF-IDF trigram",
                        "entity_type": "",
                    }
                    candidates.append(candidate)
                    tfidf_candidates[term.casefold()] = candidate

            entity_counts = Counter()
            entity_display = {}
            entity_types = {}
            for entity, entity_type in named_entities[row_index]:
                key = entity.casefold()
                entity_counts[key] += 1
                entity_display.setdefault(key, entity)
                entity_types.setdefault(key, set()).add(entity_type)

            entity_only = []
            selected_entity_keys = set(
                sorted(
                    entity_counts,
                    key=lambda key: (-entity_counts[key], key),
                )[: args.entity_top_n]
            )
            for key, occurrence_count in entity_counts.items():
                entity_type = "; ".join(sorted(entity_types[key]))
                if key in tfidf_candidates:
                    candidate = tfidf_candidates[key]
                    candidate["term"] = entity_display[key]
                    candidate["occurrences"] = occurrence_count
                    candidate["source"] = f"{candidate['source']} + NER"
                    candidate["entity_type"] = entity_type
                elif key in selected_entity_keys:
                    entity_only.append(
                        {
                            "term": entity_display[key],
                            "score": all_scores.get(key),
                            "occurrences": occurrence_count,
                            "source": "NER",
                            "entity_type": entity_type,
                        }
                    )

            entity_only.sort(
                key=lambda candidate: (
                    -candidate["occurrences"],
                    candidate["term"].casefold(),
                )
            )
            candidates.extend(entity_only)

            for rank, candidate in enumerate(candidates, start=1):
                score = candidate["score"]
                generated_term = candidate["term"]
                previous_review = review_values.get(
                    candidate_key(path.name, generated_term),
                    {},
                )
                writer.writerow(
                    [
                        path.name,
                        previous_review.get("candidate_term", generated_term),
                        f"{score:.6f}" if score is not None else "",
                        candidate["occurrences"],
                        rank,
                        candidate["source"],
                        candidate["entity_type"],
                        previous_review.get("Keep", False),
                        generated_term,
                    ]
                )

    print(
        f"Wrote {output} ({len(documents)} documents, up to {args.top_n} "
        f"TF-IDF terms, {args.trigram_top_n} additional trigrams, and "
        f"{args.entity_top_n} additional named entities per document)."
    )


if __name__ == "__main__":
    main()
