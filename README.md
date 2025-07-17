# EDI Parser

## Installation

To use the EDI Parser, you'll need to have Python 3 installed on your system. You can download the latest version of Python from the official website: [https://www.python.org/downloads/](https://www.python.org/downloads/).

Once you have Python installed, you can install the required dependencies by running the following command in your terminal or command prompt:

```
pip install tkinter openpyxl nuitka
```

## Usage

To run the EDI Parser, simply execute the `edi_parser_main.py` script:

```
python edi_parser_main.py
```

This will launch the main application window, where you can load and parse EDI files.

To compile it into .exe run:

```
py -3.12 -m nuitka --standalone --onefile --lto=yes --jobs=4 --windows-console-mode=disable --assume-yes-for-downloads --plugin-enable=anti-bloat --plugin-enable=tk-inter --python-flag=-O --nofollow-import-to=*.test,*.tests,*.unittest,*.mocks edi_parser_main.py
```

## API

The EDI Parser consists of the following Python modules:

- `edi_parser_main.py`: The main entry point for the application, which handles file selection and switching between the different EDI parsers.
- `edi_parser_trwkob.py`: The parser for TRWKOB EDI files.
- `edi_parser_minebea.py`: The parser for MINEBEA EDI files.
- `edi_parser_cummins.py`: The parser for Cummins EDI files.
- `build_nuitka.py`: A script to build the application using the Nuitka compiler.

Each parser module provides the following functionality:

- Loading and parsing EDI files
- Displaying the parsed data in a user interface
- Exporting the delivery schedule data to an Excel file

## Contributing

If you find any issues or have suggestions for improvements, please feel free to open a new issue or submit a pull request on the project's GitHub repository.

## License

This project is licensed under the [MIT License](LICENSE).
