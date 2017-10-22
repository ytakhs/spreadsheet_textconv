# spreadsheet_textconv

Command tool to convert spreadsheet file (e.g.: .xlsx, .xls) to text.

## Usage

```bash
$ spreadsheet_textconv /file/to/spreadsheet-file
```

### with git diff
```bash
$ echo "*.xlsx diff=xlsx" >> .gitattributes
$ git config --global diff.xlsx.binary true
$ git config --global diff.xlsx.textconv path/to/spreadsheet_textconv
```

## License
MIT

## See also
This tool is ported from Golang version.
- Golang version: https://github.com/tokuhirom/git-xlsx-textconv

And there are several ports:
- Perl version: https://github.com/yappo/p5-git-xlsx-textconv.pl
- Python version: https://gist.github.com/nmz787/c43bc109db915064f188
