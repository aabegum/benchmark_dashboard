# Benchmark
This program aims to compare multiple electric distribution companies based on their performance. The program uses the data from the companies and calculates the performance based on the given criteria. The program is designed to be flexible and can be adjusted to the needs of the user.

## Installation
This program is written in Python. To run the program, you need to have Python installed on your computer. You can download Python from the official website [here](https://www.python.org/downloads/). The program is written in Python 3.11. Any version of Python that is Python 3.11 or later should work.

Install the required packages by running the following command:
```bash
pip install -r requirements.txt
```

## Adjust the configuration file
The code is aimed to not change unless there is a need to do for programming purposes. The changes on the benchmark program usually are done on the configuration file. The configuration file is located at `config.yaml`.
In the configuration file, you can adjust the input parameters, the company groups, report year and type and some variables that will matter in the output file in visual or data calculation aspects. The configuration file is well commented to help you understand what each parameter does.

## The input data
There are two primary input file that the program use.
First one is the `MASTER_FILE` which contains the data of the companies. The data is in the form of a Excel file. It should be typed cleanly and correctly. It is the most important part of the program. The program will not work if the data is not correct.

For example some key points to remember:
- The Excel file should have two sheets. The first sheet should be named `{REPORT_YEAR}_Total_Veriler` where `{REPORT_YEAR}` is the report year that is typed in the configuration file. This data should have the columns `APG No`, `APG İsmi`, `Birim`. And the rest of the columns should be the companies that are defined in the configuration file. The order does not matter as they will be shuffled for the output file anyway.
- The second sheet should be named `pptx_layout`. It contains the layout of the presentation template. It should have the columns `APG Kodu`, `Sayfa`, `Left`, `Top`, `Width`, `Height`, `Grafik_tipi`, `Bulgu?`. The program will use this data to place the graphics in the presentation template.
- Grafik tipi is the type of the graphic that will be placed in the presentation. The program will use the corresponding function that is assigned for the grafik tipi.
- Bulgu? is a boolean value that indicates if the filtered mean of the data should be placed in the presentation or not corresponding to the Bulgu text boxes in the presentation template.
- APG column should be sorted correctly and the first 2 sheets should have the same APG No order. The APG No that ends with EK are exception and can be placed at the end.
- APG No should be in the form of `A1.1` where A1 represents the APG category and A1.1 represents the APG subcategory. The program processes the data based on this format. It won't work correctly if the APG No is not in this format.
- The data should be clean and without any extra rows or columns.
- The data columns should include the companies that are in the configuration file. If a company is not in the configuration file, the program will not work.
- The data of companies should not include any string value. The data should be in the form of numbers.
- The data may have integer or float values. The program will work with both types of data.
- However, it's your responsibility to round the floating point numbers to the desired decimal point. The program will not round the numbers for you.
- The percentage Birim are exceptions, they are formatted according to the `DEFAULT_DECIMAL_DIGITS` parameter in the configuration file.
- The page numbers should be in the range of the presentation template file. If the page numbers are not in the range, the program will not work.

## Running the program
To run the program, you need to run the `__main__.py` file. You can run the program by running the following command:
```bash
python __main__.py
```
The program will inform about the progress and the graphics or the presentations as they are saved. The program will run for each company group, you can see the reports and graphics under the `Main_Directory` folder with their respective company group names.
