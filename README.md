# PowerQueryViewer

The `power_query_extractor.py` script is a powerful Python utility designed for data analysts, developers, and anyone involved in data transformation and analysis using Excel's Power Query feature. It automates the extraction and decoding of Power Query M code from Excel files (.xlsx, .xlsm, .xlsb), bypassing the need to manually open and navigate through Excel documents.

## Features

- **Supports Multiple Excel Formats**: Works seamlessly with .xlsx, .xlsm, and .xlsb file formats.
- **Efficient Extraction and Decoding**: Automatically navigates through Excel's internal ZIP and XML structures to extract and decode Power Query M code, specifically targeting content within `Formulas/Section1.m`.
- **Command-Line Interface**: Designed for ease of use, enabling users to quickly extract Power Query M code through a simple command-line instruction.

## Getting Started

### Prerequisites

Ensure you have Python 3.x installed on your system. If you don't have Python installed, you can download it from the [official Python website](https://www.python.org/downloads/).

### Installation

1. Clone the repository to your local machine:
git clone https://github.com/jamesdesantiago/PowerQueryViewer.git


2. Navigate to the cloned repository directory.
3. There are no external dependencies required to run the script as it uses standard Python libraries.

## Usage

Run the script from the command line by navigating to the directory containing `power_query_extractor.py` and executing:

Replace `path_to_your_excel_file.xlsx` with the actual path to the Excel file you wish to extract Power Query M code from.

## Contributing

Contributions are what make the open-source community such an amazing place to learn, inspire, and create. Any contributions you make are **greatly appreciated**.

1. Fork the Project
2. Create your Feature Branch (`git checkout -b feature/AmazingFeature`)
3. Commit your Changes (`git commit -m 'Add some AmazingFeature'`)
4. Push to the Branch (`git push origin feature/AmazingFeature`)
5. Open a Pull Request

## License

Distributed under the Apache 2.0 License. See `LICENSE` for more information.

## Contact

James De Santiago - james.desantiago@outlook.com

Project Link: [https://github.com/jamesdesantiago/PowerQueryViewer](https://github.com/jamesdesantiago/PowerQueryViewer)

