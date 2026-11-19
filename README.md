Best practices for Indexing that this software just doesn't do yet (these first two sections are largely cribbed from the wikipedia article on indexing).

- Ensuring each topic/section includes a variety of relevant index entries; use two or three entries per topic
- Understanding the audience and understand what kind of index entries they're likely to look for
- Use the same form throughout (singular vs. plural, capitalisation, etc.), using standard indexing conventions
- Include cross-references for terms people are likely to look up but that are not actually in the text. (I have no solution for this yet from within the text.  Maybe there's a database or corpus somewhere I can refer to that would let me find related terms in the same vector space as our keywords?

Indexing pitfalls that this program embraces wholeheartedly and need to be fixed (my thoughts for solutions in parens);

- Significant topics with no index entries at all (maybe switch from TF-IDF to TextRank?  TF-IDF, even backed up by the Brown corpus, is not doing a good job)
- Indexing 'mere mentions': "But John Major was no Winston Churchill..." indexed under 'Churchill, Winston' (this is what the reviewing interface is for, probably best not to fix this with IR)
- References to discussions of a single topic scattered among several main headings: 'Cats, 50–62'; 'Felidae, 175–183' (TODO: maybe the review interface could have an easy way to demote a keyword to a cross-reference or sub-heading? Like if just before "Keep" we had a "Type" column with a drop-down menu for keyword, cross-reference, or sub-heading? Then maybe the UI needs a "re-sort" button that moves these rows together in the data file?)
- Confusing similar names: Henry V of England, Henri V of France (I think this is handled by NLTK's NER methods)

Then for an overview of how an Index works in Word, see:

     https://support.microsoft.com/en-us/word/create-and-update-an-index


here's an interesting previous solution to the first part: https://github.com/karzmei/Book-index


To generate a list of candidate terms:

`python3 ./suggest_index_terms.py`

The generator uses the NLTK Brown corpus as background documents for IDF, considers one-, two-, and three-word phrases, and adds up to 10 extra trigram candidates per chapter beyond the main TF-IDF list. If Brown has not been downloaded, install it with `python3 -m nltk.downloader brown`.

To review candidates in a desktop interface:

`python3 ./review_index_candidates.py`

The app opens `index_candidates.csv` by default. You can also pass another CSV, such as `index_candidates_with_ner.csv`. Select a row to see its extracted DOCX text with the term highlighted, check **Keep** for entries to retain, and save your choices.

The app requires PySide6 and python-docx:

`python3 -m pip install PySide6 python-docx`

The preview shows extracted text, not Word's page layout or pagination.

After reviewing, open the book in Word and run the VBA script:

     MarkReviewedIndexTerms.bas

It creates XE fields only for rows whose `Keep` value is `True`, in the Word document named on that row.
