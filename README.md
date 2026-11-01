# vcard2excel

A small Go command-line tool that converts a vCard (`.vcf`) contact export into an Excel (`.xlsx`) spreadsheet.

I wrote it to move phone contacts exported from iOS into a spreadsheet that is easy to filter, edit and share.

## How it works

- Reads `vcard.vcf` line by line.
- Picks up each contact's name (`FN:`) and phone number (`TEL`).
- Writes one row per contact to `output.xlsx` using [excelize](https://github.com/xuri/excelize).

## Requirements

- Go 1.20+

## Usage

```bash
git clone https://github.com/bunyadahmadli/vcard2excel.git
cd vcard2excel
go run .
```

Put your export next to `main.go` as `vcard.vcf` (a sample file is included). The result is written to `output.xlsx`.

## Limitations

- Only the full name and the first phone number of each contact are exported.
- Input and output file names are fixed in `main.go`.

## License

[MIT](LICENSE)
