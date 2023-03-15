# md2shunn

Converts Markdown (.md) files into Proper Manuscript Format for fiction writers
(a.k.a. Shunn Format, https://www.shunn.net/format/) as a Word (.docx) file.

## Example usage

```
md2shunn --input draft01.md --author 'Edgar Allan Poe' --title 'The Tell-Tale Heart'
# Outputs to draft01.docx
```

## Features

* Automatic conversion of straight to smart quotes and apostrophes
* Automatic word count
* Automatic conversion of full title and author to truncated versions for header
* Supports both [Modern](https://www.shunn.net/format/story/) and [Classic](https://www.shunn.net/format/classic/) formats via `--format` flag