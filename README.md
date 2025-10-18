# Automated Ratio Calculator (ARC)
An Arelle Plugin for Automated Ratio Calculation

Automated Ratio Calculator (ARC) is a plugin for Arelle that provides enhanced table detection and automated financial ratio calculation for German and IFRS filings.

## Description

This plugin extends Arelle's functionality by automatically detecting financial tables in filings and calculating financial ratios (IFRS & US-GAAP).

## Use Case

With ARC, in principle, you can download any iXBRL and calculate the financial ratios without delay or reliance on third party providers. This may lead to faster and better financial decisions, because the original source data is used the minute it is made public. Just download iXBRL data from companies' investor relations websites, SEC's EDGAR or any other official Government Register (e.g. Unternehmensregister for Germany) and process the data through ARC. The output will be a detailed excel sheet with ratios and traceability how the ratios were calulated.

## Features

- Automated detection and extraction of tables from financial filings
- Calculation of financial ratios and output to summary excel file
- Support for German filings
- Support for IFRS and US-GAAP accounting standards
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

## Financial Analysis Disclaimer

The ARC financial ratios and analyses are calculated based on the data extracted from XBRL filings. Please note the following important considerations:

1. The calculations are automated and should be verified against the original financial statements.

2. Different accounting standards and company-specific reporting practices may affect the comparability of ratios across different companies or periods.

3. This analysis is provided for informational purposes only and should not be considered as financial advice or a recommendation for any investment decision.

4. Users should conduct their own due diligence and consult with qualified financial advisors before making any financial decisions.

5. The accuracy of these ratios depends on the accuracy and completeness of the underlying XBRL data in the filing.

6. The creators and contributors of this software plugin are not liable for any errors, omissions, or any consequences arising from the use of this analysis tool. Use of this software is at your own risk.

---

*Note: This plugin is designed to work with Arelle, an open source XBRL platform. For more information about Arelle, visit [Arelle's official website](https://arelle.org/).*
