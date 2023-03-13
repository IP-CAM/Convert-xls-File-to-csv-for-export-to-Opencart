## XLS to CSV Converter Application

![License](https://img.shields.io/github/license/aschmelyun/larametrics.svg?style=flat-square)

This is a Python application that uses the Qt library to provide a graphical user interface for converting XLS files to CSV. The purpose of this application is to allow users to easily convert their XLS files into a format that can be used to update product prices in OpenCart.

![screen](https://github.com/vstaran/convert-to-csv/blob/main/screen.png?raw=true)


## Installation

1. Clone this repository to your local machine.
2. Run the application using the following command:

```bash
pyuic5 form.ui -o form.py
python main.py
```


## Build

```bash
pyinstaller --noconsole --onefile main.py
```


## Usage

1. Launch the application by running the main.py file.
2. Click the "Choose File" button to select the XLS file you want to convert.
3. Select the output directory where you want to save the converted CSV file.
4. Click the "Convert" button to convert the XLS file to CSV.
5. Once the conversion is complete, you can use the resulting CSV file to update product prices in OpenCart.


## License

This application is licensed under the [MIT License](https://opensource.org/license/mit/). Feel free to use, modify, and distribute the code as needed.
