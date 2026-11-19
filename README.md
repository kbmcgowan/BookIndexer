## Install & Run

This works on my mac. Your python3 binary might be called something else.

### Install

Change directory to the root of this project and execute:

```sh
python3 -m venv .BookIndexer-venv
source .BookIndexer-venv/bin/activate
python -m pip install --upgrade pip
python -m pip install -r requirements.txt
```

If you don't have it already, you need to download the nltk data using python

```python
import nltk
nltk.download('brown')
```

subsequently you can run the indexer with:

```sh
source .BookIndexer-venv/bin/activate
python suggest_index_terms.py
```

or you can review candidates with:

```sh
source .BookIndexer-venv/bin/activate
python review_index_candidates.py &
```

`deactivate` will exit the venv

Best practices for Indexing that this software just doesn't do yet (these first two sections were initially cribbed from the [wikipedia article on indexing](https://en.wikipedia.org/wiki/Index_(publishing))).

- Ensuring each topic/section includes a variety of relevant index entries; use two or three entries per topic
- Understanding the audience and understand what kind of index entries they're likely to look for
- Use the same form throughout (singular vs. plural, capitalisation, etc.), using standard indexing conventions
- Include cross-references for terms people are likely to look up but that are not actually in the text. (I have no solution for this yet from within the text.  Maybe there's a database or corpus somewhere I can refer to that would let me find related terms in the same vector space as our keywords?

Indexing pitfalls that this program embraces wholeheartedly and need to be fixed (my thoughts for solutions in parens);

- Significant topics with no index entries at all (maybe switch from TF-IDF to TextRank?  TF-IDF, even backed up by the Brown corpus, is not doing a good job)
- Indexing 'mere mentions': "But John Major was no Winston Churchill..." indexed under 'Churchill, Winston' (this is what the reviewing interface is for, probably best not to fix this with IR)
- References to discussions of a single topic scattered among several main headings: 'Cats, 50–62'; 'Felidae, 175–183' (TODO: maybe the review interface could have an easy way to demote a keyword to a cross-reference or sub-heading? Like if just before "Keep" we had a "Type" column with a drop-down menu for keyword, cross-reference, or sub-heading? Then maybe the UI needs a "re-sort" button that moves these rows together in the data file?)
- Confusing similar names: Henry V of England, Henri V of France (I think this is handled by NLTK's NER methods)
- TODO: make sure any languages mentioned in the text are (a) identified by NER and (b) labeled 'language' in the index (e.g. !Xóõ language and Yoruba language)
- TODO: harder, figure out how to do these kind of ... clarification? annotations like the '(Maya)' in: Yucatec (Maya) language  (handle in review UI? what if the Keyword column were editable?  https://www.pythonguis.com/faq/editing-pyside6-tableview/ this could be very powerful but makes the csv file substantially more valuable/less reproducible)
- add TextRank as a concept-identifying tool alongside TF-IDF, NER, and the current trigram hack


Then for an overview of how an Index works in Word, see:

     https://support.microsoft.com/en-us/word/create-and-update-an-index


here's an interesting previous solution to the first part: https://github.com/karzmei/Book-index


To generate a list of candidate terms:

`python3 ./suggest_index_terms.py`

The generator uses the NLTK Brown corpus as background documents for IDF, considers one-, two-, and three-word phrases, and adds up to 10 extra trigram candidates per chapter beyond the main TF-IDF list. If Brown has not been downloaded, install it with `python3 -m nltk.downloader brown`.

The shared TOML preferences are in `preferences.toml`. By default, DOCX files are read from `./book/` and candidate CSVs are read/written in `./index/`, relative to the preferences file. Put the chapter DOCX files in `book/`, or change `book_path` in the reviewer’s Preferences dialog. The generator also accepts `--directory` and `--output` overrides.

To review candidates in a desktop interface:

`python3 ./review_index_candidates.py`

The app opens `index_candidates.csv` by default. You can also pass another CSV, such as `index_candidates_with_ner.csv`. Select a row to see its extracted DOCX text with the term highlighted, edit the candidate term, check **Keep** for entries to retain, and save your choices. Term edits are preserved when the generator is rerun, matched by the hidden `source_candidate_term` column.

Use the **Preferences...** button to change the book and index folders and select which candidate columns to hide. The `source` column is hidden by default; `source_candidate_term` is always hidden because it is internal matching metadata.

The app requires PySide6 and python-docx:

`python3 -m pip install PySide6 python-docx`

The preview shows extracted text, not Word's page layout or pagination.

After reviewing, open the book in Word and run the VBA script:

     MarkReviewedIndexTerms.bas

It creates XE fields only for rows whose `Keep` value is `True`, in the Word document named on that row.
