# AutomatedRatioCalculator
An Arelle Plugin for Automated Ratio Calculation

Automated Ratio Calculator (ARC) is a plugin for Arelle that provides enhanced table detection and automated financial ratio calculation for German and IFRS filings.

## Description

This plugin extends Arelle's functionality by automatically detecting tables in financial statements and calculating important financial ratios. It specializes in processing both German accounting standards and IFRS (International Financial Reporting Standards) filings.

## Features

- Automated detection and extraction of tables from financial filings
- Calculation of standard financial ratios
- Support for both German accounting standards and IFRS
- Seamless integration with Arelle

## Installation

1. Copy files from `AutomatedRatioCalculator/tree/main/plugin/ARC` to your Arelle installation directory
   - Example path: `C:\Program Files\Arelle\plugin\ARC\__init__.py`

2. Unpack files from `AutomatedRatioCalculator/tree/main/lib/lib.rar` to the Arelle/lib/ directory
   - These files are necessary because ARC uses Python libraries that are not included with Arelle by default

## Usage

1. Start Arelle.exe
2. Go to Help → Manage plug-ins
3. Load and enable the Automated Ratio Calculator plugin by clicking on browse then select the `__init__.py` file on your drive
4. Restart Arelle for the changes to take effect
5. Check and save ARC settings: Tools → Configure Automated Ratio Calculator Plugin. Change the output folders to your liking
6. Load your financial filing (web resource or .zip-file)
7. Review the ratios in the output folders

## Requirements

- Arelle XBRL platform
- Python libraries included in the lib.rar file

## License

This project is licensed under the Apache License 2.0 - see the LICENSE file for details.

## Contributing

Contributions are welcome! Please feel free to submit a Pull Request.

## Support

If you encounter any problems or have any questions, please open an issue on this repository.

---

*Note: This plugin is designed to work with Arelle, an open source XBRL platform. For more information about Arelle, visit [Arelle's official website](https://arelle.org/).*
