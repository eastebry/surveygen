# Multilingual survey

Excel stores the wording and translations. A question list controls the order
and answer types. LaTeX handles the page layout, and Docker runs everything to
create a separate PDF for each language.

## Create the PDFs

Install and open Docker Desktop. Open Terminal in this project folder and run:

```sh
./build.sh
```

Find the finished surveys in **output/pdf**. The first run takes longer while
it downloads the tools it needs.

## Change the wording

Open **translations.xlsx** in Excel. On the **Translations** sheet, edit the
English (`en`), Chinese (`zh-Hans`), or Arabic (`ar`) columns. These include the
questions, answers, title, agency name, and cover text.

Keep the labels in the **key** column unchanged, fill in every translation, and
preserve `{page}` and `{pages}` in the page-number wording. Save and rerun the
command above. Layout or logo changes need someone familiar with LaTeX.

## Add or rearrange questions

1. Open **survey.json** in a text editor. Copy a question of the same type:
   `single` (choose one), `multiple` (choose several), or `matrix` (rating table).
2. Give it a unique `id` and new text labels. Add matching **key** rows in Excel
   with the question, answers, and all translations. A technical colleague can
   help preserve the question file's punctuation and structure.
3. Place it where you want it in the list. Numbering updates automatically.
   Save both files and regenerate the PDFs.

## Add a language

Add an Excel column, such as `fr` for French, and translate every row. Ask a
technical colleague to register that code in **scripts/build.py**, add its
language settings in **tex/languages**, and include any needed fonts in the
**Dockerfile**. Then regenerate and review the PDFs.

Always check every page after changes. Questions must fit within two or three
pages, plus the cover. If generation reports an error, existing PDFs have **not**
been updated.
