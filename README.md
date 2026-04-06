# Corrupt Extractor for Microsoft Office

Extracts text and data from corrupt DOCX, XLSX, and PPTX files. Advanced mode can fix zip structure, recover embedded images, and edit corrupt XML directly.

## Screenshots

Visit the [SourceForge project page](https://sourceforge.net/projects/crrptoffcxtrctr/) to view screenshots.

> **Tip:** If you have screenshots to contribute, open a PR adding them to a `screenshots/` folder!

**Language:** Delphi 7  
**License:** MIT

## Features

- Extracts text from corrupt Office 2007+ files (DOCX, XLSX, PPTX)
- Advanced mode fixes zip archive structure
- Recovers embedded images from damaged documents
- Direct XML editing for manual repair
- Handles files that Office cannot open at all

## System Requirements

- Windows XP or later
- Delphi 7 (for original build) or Free Pascal / Lazarus (free alternative)

## Installation & Usage

### Building from Source (Delphi 7)

1. Open the `.dpr` project file in Delphi 7
2. Press **F9** to compile and run

### Building with Free Pascal (free alternative)

```bash
sudo apt-get install fpc    # Linux
# or download from https://www.freepascal.org/
fpc -Sd src/*.pas
```

### Using a Pre-built Release

Download the latest release from the [Releases](../../releases) page.

## Origin

This project was originally hosted on SourceForge and has been migrated to GitHub for easier access and collaboration.

- **SourceForge:** [crrptoffcxtrctr](https://sourceforge.net/projects/crrptoffcxtrctr/)
- **Migrated with:** [SF2GH Migrator](https://github.com/socrtwo/sf-to-github)

## Contributing

Contributions are welcome! Feel free to:

1. Fork this repository
2. Create a feature branch (`git checkout -b my-feature`)
3. Commit your changes (`git commit -m "Add my feature"`)
4. Push to the branch (`git push origin my-feature`)
5. Open a Pull Request

## License

MIT License — see [LICENSE](LICENSE) for details.
