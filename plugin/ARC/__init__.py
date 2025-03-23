"""
Automated Ratio Calculator Plugin and enhanced table detection for IFRS/ESEF and German filings
"""
from arelle import XbrlConst
from arelle.ModelDtsObject import ModelConcept
from arelle.XmlValidateConst import VALID
from arelle.ModelXbrl import ModelXbrl
from arelle import TableStructure  # Needed for table rendering
from typing import Dict, List, Tuple, Optional, Set
import regex as re
import os
from collections import defaultdict
from datetime import datetime, timedelta
import os
import traceback
import pandas as pd
from openpyxl.styles import PatternFill, Font, Border, Side, Alignment
from openpyxl.utils import get_column_letter
from collections import Counter, defaultdict
# from threading import Lock
from time import perf_counter
import json
import tkinter as tk
from tkinter import filedialog
from tkinter import ttk, messagebox
import os

class Logger:
    def __init__(self):
        self.log_path = None
        self.enabled = False
        self._initialized = False

    # Hardcoded        
    # def initialize(self, model_xbrl):
    #     try:
    #         output_dir = r"C:\Data\venv_ArellePlugin_Release\sec_tables_output"
    #         os.makedirs(output_dir, exist_ok=True)
            
    #         filename = os.path.basename(model_xbrl.uri) if model_xbrl.uri else "unknown_filing"
    #         base_filename = os.path.splitext(filename)[0]
    #         self.log_path = os.path.join(output_dir, f"{base_filename}_table_detection.log")
            
    #         with open(self.log_path, 'w', encoding='utf-8') as f:
    #             f.write(f"=== Log Started at {datetime.now()} ===\n")
                
    #     except Exception as e:
    #         print(f"Error initializing logger: {str(e)}")

    # config file
    def initialize(self, model_xbrl):
        """Initialize logger with robust state tracking"""
        try:
            print("\n=== Logger Initialization ===")
            
            if self._initialized:
                print("Logger already initialized")
                return
                
            # Check config handler
            if not hasattr(model_xbrl, '_config_handler'):
                print("No config handler available - logging disabled")
                self.enabled = False
                return
                
            # Get and verify config
            config_handler = model_xbrl._config_handler
            if not config_handler.config:
                print("No config available - using defaults")
                self.enabled = True  # Default to enabled
            else:
                self.enabled = config_handler.is_logging_enabled()
            
            print(f"Logging enabled status: {self.enabled}")
            
            if not self.enabled:
                print("Logging disabled - skipping file setup")
                self._initialized = True
                return
                
            # Get and verify export path
            export_path = config_handler.get_export_path(model_xbrl)
            if not export_path:
                print("No valid export path - logging disabled")
                self.enabled = False
                self._initialized = True
                return
                
            print(f"Using export path: {export_path}")
            
            # Setup log file
            try:
                filename = os.path.basename(model_xbrl.uri) if model_xbrl.uri else "unknown_filing"
                base_filename = os.path.splitext(filename)[0]
                self.log_path = os.path.join(export_path, f"{base_filename}_table_detection.log")
                
                # Ensure directory exists
                os.makedirs(export_path, exist_ok=True)
                
                # Create/verify log file
                with open(self.log_path, 'w', encoding='utf-8') as f:
                    f.write(f"=== Log Started at {datetime.now()} ===\n")
                    f.write(f"Export Path: {export_path}\n")
                    f.write(f"Config Status: Enabled={self.enabled}\n")
                    f.write(f"Config: {json.dumps(config_handler.config, indent=2)}\n")
                
                if os.path.exists(self.log_path):
                    print(f"Successfully created log file: {self.log_path}")
                else:
                    print(f"Failed to create log file: {self.log_path}")
                    self.enabled = False
                    
            except Exception as e:
                print(f"Error setting up log file: {str(e)}")
                print(traceback.format_exc())
                self.enabled = False
                self.log_path = None
                
            self._initialized = True
            
        except Exception as e:
            print(f"Error in logger initialization: {str(e)}")
            print(traceback.format_exc())
            self.enabled = False
            self._initialized = True
            
        finally:
            print(f"\nLogger initialization complete:")
            print(f"Enabled: {self.enabled}")
            print(f"Path: {self.log_path}")
            print(f"Initialized: {self._initialized}")


    def log(self, message: str, level: str = "INFO", include_trace: bool = False):
        """Enhanced logging with better error handling and debugging"""
        try:
            if not self.enabled:
                print(f"[Log Skipped - Disabled] {message}")
                return
                
            if not self.log_path:
                print(f"[Log Skipped - No Path] {message}")
                return
                
            timestamp = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
            try:
                with open(self.log_path, "a", encoding='utf-8') as f:
                    f.write(f"[{timestamp}] {level}: {message}\n")
                    if include_trace:
                        f.write(f"Stack trace:\n{traceback.format_exc()}\n")
                print(f"[Log Written] {self.log_path}: {message}")
            except IOError as e:
                print(f"Failed to write to log file {self.log_path}: {str(e)}")
                print(f"Attempted to write: {message}")
                        
        except Exception as e:
            print(f"Logging error: {str(e)}")
            print(f"Failed to write: {message}")
    
    # old    
    # def log(self, message: str, level: str = "INFO", include_trace: bool = False):
    #     """Enhanced logging with optional stack trace"""
    #     try:
    #         if self.log_path:
    #             timestamp = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
    #             with open(self.log_path, "a", encoding='utf-8') as f:
    #                 f.write(f"[{timestamp}] {level}: {message}\n")
    #                 if include_trace:
    #                     f.write(f"Stack trace:\n{traceback.format_exc()}\n")
    #     except Exception as e:
    #         print(f"Logging error: {str(e)}")

# Update the log_debug wrapper function to handle all parameters
# def log_debug(message: str, level: str = "DEBUG", include_trace: bool = False):
#     logger.log(message, level, include_trace)

def log_debug(message: str, level: str = "DEBUG", include_trace: bool = False):
    """Enhanced debug logging with file and console output"""
    # Always print to console
    print(f"[{level}] {message}")
    
    # Only try file logging if logger is properly initialized
    global logger
    if logger and hasattr(logger, 'enabled') and logger.enabled and hasattr(logger, 'log'):
        logger.log(message, level, include_trace)

def write_summary(message):
    """Enhanced summary logging"""
    global logger
    if logger and hasattr(logger, 'enabled') and logger.enabled:
        logger.log(message, "SUMMARY")
    else:
        print(f"[SUMMARY] {message}")


# Initialize global logger
logger = Logger()

            

class StatementTypeDetector:
    """Detects financial statement types using multiple criteria and probability scoring"""
    
    def __init__(self):
        self.PARENTHETICAL_PATTERN = r"pa?r[ae]ne?th\w?[aei]+\w?t?h?i?c"
        self.MIN_STATEMENT_SIZE = 15
        self.patterns = self._initialize_patterns()
        self.mandatory_facts = self._initialize_mandatory_facts()
        
    def _initialize_patterns(self) -> Dict[str, Dict[str, List[str]]]:
        """Initialize comprehensive pattern sets for statement detection"""
        return {
            'BS': {
                'primary': [
                    r'.*?(statement.*?financial.*?position|balance.*?sheet|bilanz|vermögenslage|finanzlage)',
                    r'.*?(consolidated.*?statement.*?financial.*?position)',
                    r'.*balance\s+sheets?.*',
                    r'.*statements?\s+of\s+financial\s+position.*',
                    r'.*bilanz.*',
                    r'.*konzernbilanz.*',
                ],
                'secondary': [
                    r'.*financial\s+position.*',
                    r'.*balance.*',
                    r'.*aktiva.*passiva.*',
                ],
                'indicators': ['assets', 'liabilities', 'vermögen', 'schulden']
            },
            'IS': {
                'primary': [
                    r'.*?(statement.*?(profit|loss|income)|income.*?statement|gewinn.*?verlust)',
                    r'.*statements?\s+of\s+income.*',
                    r'.*income\s+statements?.*',
                    r'.*gewinn.*verlust.*rechnung.*',
                    r'.*konzern.*gewinn.*verlust.*',
                ],
                'secondary': [
                    r'.*profit.*loss.*',
                    r'.*income.*',
                    r'.*earnings.*',
                    r'.*\bPnl\b.*',
                    r'.*\bkonzern.*GuV\b.*',    # special case for E.ON     
                ],
                'indicators': ['revenue', 'profit', 'loss', 'umsatz', 'ertrag'],
                'negative': ['comprehensive']  # Must not contain these
            },
            'CI': {
                'primary': [
                    r'.*?(comprehensive.*?income|gesamtergebnis)',
                    r'.*comprehensive\s+income.*',
                    r'.*gesamtergebnis.*rechnung.*',
                    r'.*konzern.*gesamtergebnis.*'                 
                ],
                'secondary': [
                    r'.*total\s+comprehensive.*',
                    r'.*other\s+comprehensive.*',
                    r'.*\bci[\s-]*tabelle\b.*',
                    r'.*\bconsolidated\s*ci\b.*',                   
                    r'.*\bOci\b.*', # special case just for Airbus
                    r'.*\bci\b.*',
                    r'.*gesamtergebis.*rechnung.*', # Special Case for Typo in Infineon Report
                ],
                'indicators': ['comprehensive', 'gesamtergebnis']
            },
            'CF': {
                'primary': [
                    r'.*?(cash.*?flow|kapitalfluss|cashflow|geldfluss)',
                    r'.*cash\s+flows?.*',
                    r'.*kapitalfluss.*rechnung.*',
                    r'.*konzern.*kapitalfluss.*',
                ],
                'secondary': [
                    r'.*cash\s+flow.*statement.*',
                    r'.*statements?\s+of\s+cash\s+flows?.*',
                    r'.*\bcf[\s-]*tabelle\b.*',
                    r'.*\bconsolidated\s*cf\b.*',    
                    r'.*\bkonzern.*KFR\b.*',    # special case for E.ON: Kapitalflussrechnung          
                ],
                'indicators': ['cash', 'flows', 'kapitalfluss']
            },
            'EQ': {
                'primary': [
                    r'.*?(changes.*?equity|eigenkapitalveränderung)'
                    r".*shareholders?['']\s*equity.*",
                    r'.*statements?\s+of\s+changes\s+in\s+equity.*',
                    r'.*eigenkapitalveränderung.*',
                    r'.*konzern.*eigenkapital.*',
                ],
                'secondary': [
                    r'.*changes?\s+in\s+equity.*',
                    r'.*stockholders.*equity.*',
                    r'.*\bce[\s-]*tabelle\b.*',
                    r'.*\bconsolidated\s*ce\b.*', 
                    r'.*\bchangesinequity\b.*',
                    r'.*\bkonzern.*EK\b.*',    # special case for E.ON     
                    r'(?i)\b(veränderung?|entwicklung)\s+des\s+eigenkapitals\b',    # special case for RWE and Rheinmetall                    
                ],
                'indicators': ['equity', 'eigenkapital', 'shareholders']
            }
        }

    def _initialize_mandatory_facts(self) -> Dict[str, Dict[str, List[str]]]:
        """Initialize mandatory facts for content verification"""
        return {
            'BS': {
                'ifrs': [
                    'ifrs-full:Assets',
                    'ifrs-full:Liabilities',
                    'ifrs-full:Equity'
                ],
                'us-gaap': [
                    'us-gaap:Assets',
                    'us-gaap:Liabilities',
                    'us-gaap:StockholdersEquity'
                ]
            },
            'IS': {
                'ifrs': [
                    'ifrs-full:Revenue',
                    'ifrs-full:ProfitLoss'
                ],
                'us-gaap': [
                    'us-gaap:Revenues',
                    'us-gaap:NetIncomeLoss'
                ]
            },
            'CI': {
                'ifrs': [
                    'ifrs-full:ComprehensiveIncome',
                    'ifrs-full:OtherComprehensiveIncome'
                ],
                'us-gaap': [
                    'us-gaap:ComprehensiveIncomeNetOfTax',
                    'us-gaap:OtherComprehensiveIncomeLossNetOfTax'
                ]
            },
            'CF': {
                'ifrs': [
                    'ifrs-full:CashFlowsFromUsedInOperatingActivities',
                    'ifrs-full:CashFlowsFromUsedInInvestingActivities',
                    'ifrs-full:CashFlowsFromUsedInFinancingActivities'
                ],
                'us-gaap': [
                    'us-gaap:NetCashProvidedByUsedInOperatingActivities',
                    'us-gaap:NetCashProvidedByUsedInInvestingActivities',
                    'us-gaap:NetCashProvidedByUsedInFinancingActivities'
                ]
            },
            'EQ': {
                'ifrs': [
                    'ifrs-full:ChangesInEquity',
                    'ifrs-full:IssuedCapital'
                ],
                'us-gaap': [
                    'us-gaap:StockholdersEquity',
                    'us-gaap:CommonStock'
                ]
            }
        }

    def detect_statement_type(self, definition: str, element_count: int = 0) -> Tuple[str, float]:
        """Detect statement type using multiple criteria and probability scoring"""
        if not definition:
            return None, 0.0
            
        definition_lower = definition.lower()
        is_parenthetical = bool(re.search(self.PARENTHETICAL_PATTERN, definition_lower))
        
        # Initialize probabilities for each type
        probabilities = defaultdict(float)
        
        # Base probability adjustments
        is_statement = bool(re.match(r".* - statement - ", definition_lower))
        has_details = not bool(re.match(r"(?!.*details)", definition_lower))
        
        base_prob = 0.5 if is_statement else 0.3
        if has_details:
            base_prob *= 0.8
            
        # Size-based probability adjustment
        size_factor = min(1.0, element_count / self.MIN_STATEMENT_SIZE) if element_count > 0 else 0.5
        
        # Check each statement type
        for stmt_type, patterns in self.patterns.items():
            prob = base_prob
            primary_matched = False
            
            # Try primary patterns first (highest weight)
            # log_debug(f"\nChecking primary patterns for {stmt_type}:")
            for pattern in patterns['primary']:
                if re.search(pattern, definition_lower, re.I):
                    # log_debug(f"Matched primary pattern: {pattern}")
                    prob += 0.4
                    primary_matched = True
                    break
                    
            # Only try secondary patterns if no primary match
            if not primary_matched:
                # log_debug("No primary match, checking secondary patterns:")
                for pattern in patterns['secondary']:
                    if re.search(pattern, definition_lower, re.I):
                        # log_debug(f"Matched secondary pattern: {pattern}")
                        prob += 0.2
                        break
                        
            # Only check indicators if we had any pattern match
            if primary_matched or prob > base_prob:
                # log_debug("Checking indicators:")
                for indicator in patterns['indicators']:
                    if indicator.lower() in definition_lower:
                        # log_debug(f"Found indicator: {indicator}")
                        prob += 0.1
                        
                # Negative indicators (for exclusion)
                if 'negative' in patterns:
                    for neg in patterns['negative']:
                        if neg.lower() in definition_lower:
                            log_debug(f"Found negative indicator: {neg}")
                            prob *= 0.1
            
            # Apply size factor
            prob *= size_factor
            
            # Special handling for CI vs IS
            if stmt_type == 'IS' and 'comprehensive' in definition_lower:
                prob *= 0.1
            elif stmt_type == 'CI' and 'comprehensive' in definition_lower:
                prob += 0.3
                
            probabilities[stmt_type] = min(1.0, prob)
            log_debug(f"Final probability for {stmt_type}: {prob:.2f}")
            
        # Get highest probability type
        if probabilities:
            best_type, highest_prob = max(probabilities.items(), key=lambda x: x[1])
            log_debug(f"\nBest match: {best_type} with probability {highest_prob:.2f}")
            
            # Add parenthetical suffix if needed
            if is_parenthetical:
                best_type += 'P'
                
            # If probability too low, mark as OTH
            if highest_prob < 0.3:
                return 'OTH', highest_prob
            return best_type, highest_prob
                
        return 'OTH', 0.0

    def debug_detection(self, definition: str, element_count: int = 0):
        """Debug helper for diagnostic information"""
        if not definition:
            log_debug(f"Empty definition provided")
            return
            
        log_debug(f"\nAnalyzing definition: {definition}")
        log_debug(f"Element count: {element_count}")
        
        definition_lower = definition.lower()
        is_statement = bool(re.match(r".* - statement - ", definition_lower))
        has_details = not bool(re.match(r"(?!.*details)", definition_lower))
        is_parenthetical = bool(re.search(self.PARENTHETICAL_PATTERN, definition_lower))
        
        log_debug(f"Is statement: {is_statement}")
        log_debug(f"Has details: {has_details}")
        log_debug(f"Is parenthetical: {is_parenthetical}")
        
        stmt_type, prob = self.detect_statement_type(definition, element_count)
        log_debug(f"Detected type: {stmt_type} (probability: {prob:.2f})")

    def detect_split_statement(self, definition: str) -> Tuple[bool, str]:
        """
        Detect if a statement is part of a split statement (e.g., Aktiva/Passiva)
        Returns (is_split, section) tuple
        """
        log_debug(f"\nChecking for split statement: {definition}")
        definition_lower = definition.lower()
        
        # Check for German balance sheet splits
        if re.search(r'\baktiva\b', definition_lower):
            log_debug("Detected AKTIVA section")
            return True, 'assets'
        elif re.search(r'\bpassiva\b', definition_lower):
            log_debug("Detected PASSIVA section")
            return True, 'liabilities'
            
        # Check for English balance sheet splits
        if re.search(r'\bassets\b', definition_lower) and not re.search(r'liabilit|equity', definition_lower):
            log_debug("Detected ASSETS section")
            return True, 'assets'
        elif re.search(r'\bliabilit|equity\b', definition_lower) and not re.search(r'\bassets\b', definition_lower):
            log_debug("Detected LIABILITIES section")
            return True, 'liabilities'
        
        log_debug("No split detected")
        return False, None

class FinancialConceptMaps:
    """Maps financial concepts across taxonomies and defines their locations"""
    
    @staticmethod
    def get_concept_maps():
        """
        Returns mapping of financial concepts
        Format: {
            'concept_key': {
                'statements': ['IS', 'CF', etc],  # Where to look for this concept
                'taxonomies': {
                    'ifrs-full': ['ConceptName1', 'AlternativeName1'],
                    'us-gaap': ['ConceptName1', 'AlternativeName1'],
                    'generic': ['OtherPossibleNames']  # For checking custom taxonomies
                }
            }
        }
        """
        return {
            # Revenue concepts (Income Statement)
            'revenue': {
                'statements': ['IS'],
                'taxonomies': {
                    'ifrs-full': ['Revenue','InsuranceRevenue','RevenueFromInterest','RevenueAndOperatingIncome','ShareOfProfitLossOfAssociatesAndJointVenturesAccountedForUsingEquityMethod','RevenueFromContractsWithCustomers'], # classical companies, insurances, banks, StockExchange, Holding
                    'us-gaap': [
                        'Revenues', 
                        'SalesRevenueNet',
                        'RevenueFromContractWithCustomerExcludingAssessedTax',
                        'RevenueFromContractWithCustomer'
                    ],
                    'generic': [
                        'TotalRevenue',
                        'GroupRevenue',
                        'ConsolidatedRevenue'
                    ]
                }
            },
            
            # Asset concepts (Balance Sheet)
            'total_assets': {
                'statements': ['BS'],
                'taxonomies': {
                    'ifrs-full': ['Assets','CurrentAssetsOtherThanAssetsOrDisposalGroupsClassifiedAsHeldForSaleOrAsHeldForDistributionToOwners'],  # std, Airbus/Heidelberg Materials
                    'us-gaap': ['Assets'],
                    'generic': [
                        'TotalAssets',
                        'ConsolidatedAssets',
                        'GroupAssets'
                    ]
                }
            },
            'current_assets': {
                'statements': ['BS'],
                'taxonomies': {
                    'ifrs-full': ['CurrentAssets','CurrentAssetsOtherThanAssetsOrDisposalGroupsClassifiedAsHeldForSaleOrAsHeldForDistributionToOwners'],
                    'us-gaap': [
                        'AssetsCurrent',
                        'CurrentAssets'
                    ],
                    'generic': [
                        'TotalCurrentAssets',
                        'GroupCurrentAssets'
                    ]
                }
            },
            
            # Liability concepts (Balance Sheet)
            'total_liabilities': {
                'statements': ['BS'],
                'taxonomies': {
                    'ifrs-full': ['Liabilities'],
                    'us-gaap': ['Liabilities'],
                    'generic': [
                        'TotalLiabilities',
                        'ConsolidatedLiabilities',
                        'GroupLiabilities'
                    ]
                }
            },
            'current_liabilities': {
                'statements': ['BS'],
                'taxonomies': {
                    'ifrs-full': ['CurrentLiabilities','CurrentLiabilitiesOtherThanLiabilitiesIncludedInDisposalGroupsClassifiedAsHeldForSale'],
                    'us-gaap': [
                        'LiabilitiesCurrent',
                        'CurrentLiabilities'
                    ],
                    'generic': [
                        'TotalCurrentLiabilities',
                        'GroupCurrentLiabilities'
                    ]
                }
            },
            
            # Cash Flow concepts
            'operating_cashflow': {
                'statements': ['CF'],
                'taxonomies': {
                    'ifrs-full': [
                        'CashFlowsFromUsedInOperatingActivities',
                        'NetCashFlowsFromUsedInOperatingActivities'
                    ],
                    'us-gaap': [
                        'NetCashProvidedByUsedInOperatingActivities',
                        'CashProvidedByUsedInOperatingActivities'
                    ],
                    'generic': [
                        'OperatingCashFlow',
                        'NetOperatingCashFlow'
                    ]
                }
            },
            
            # Profit concepts (Income Statement)
            'net_income': {
                'statements': ['IS'],
                'taxonomies': {
                    'ifrs-full': [
                        'ProfitLoss',
                        'ProfitLossAttributableToOwnersOfParent'
                    ],
                    'us-gaap': [
                        'NetIncomeLoss',
                        'NetIncomeLossAvailableToCommonStockholdersBasic',
                        'ProfitLoss'
                    ],
                    'generic': [
                        'NetProfit',
                        'GroupProfit',
                        'ConsolidatedProfit',
                        'NetIncome',
                        'Beteiligungsergebnis',
                    ]
                }
            },
            
            # Cost of Sales (Income Statement)
            'cost_of_sales': {
                'statements': ['IS'],
                'taxonomies': {
                    'ifrs-full': [
                        'CostOfSales',
                        'CostOfGoodsSold'
                    ],
                    'us-gaap': [
                        'CostOfGoodsAndServicesSold',
                        'CostOfRevenue',
                        'CostOfGoodsSold'
                    ],
                    'generic': [
                        'TotalCostOfSales',
                        'GroupCostOfSales',
                        'CostOfSalesAndServices'
                    ]
                }
            },        
            'gross_profit': {
                'statements': ['IS'],
                'taxonomies': {
                    'ifrs-full': [
                        'GrossProfit',
                        'GrossProfitLoss'
                    ],
                    'us-gaap': [
                        'GrossProfit',
                        'GrossProfitLoss',
                        'RevenueMinusCostOfRevenue'
                    ],
                    'generic': [
                        'TotalGrossProfit',
                        'GrossIncome',
                        'BruttoErgebnis',
                        'OperatingGrossProfit',
                        'RohErtrag'
                    ]
                }
            },
            'total_debt': {
                'statements': ['BS'],
                'taxonomies': {
                    'ifrs-full': [
                        'Borrowings',
                        'TotalDebt',
                        # 'NonCurrentBorrowings', 
                        # 'CurrentBorrowings',
                        # 'LongTermBorrowings',      
                        # 'LongtermBorrowings',      
                        # 'LongTermBorrowings',      
                        # 'NoncurrentInterestBearingBorrowings'                        
                    ],
                    'us-gaap': [
                        'DebtCurrent',
                        # 'LongTermDebt',
                        # 'LongTermDebtNoncurrent',
                        'DebtAndCapitalLeaseObligations'
                    ],
                    'generic': [
                        'TotalDebt',
                        'TotalBorrowings',
                        # 'Verbindlichkeiten',
                        'GesamtVerbindlichkeiten',
                        # 'Finanzschulden',          
                        'GesamtFinanzschulden'                            
                    ]
                }
            },
            'total_equity': {
                'statements': ['BS'],
                'taxonomies': {
                    'ifrs-full': [
                        'Equity',
                        'TotalEquity',
                        'EquityAttributableToOwnersOfParent'
                    ],
                    'us-gaap': [
                        'StockholdersEquity',
                        'StockholdersEquityIncludingPortionAttributableToNoncontrollingInterest',
                        'PartnersCapital'
                    ],
                    'generic': [
                        'TotalEquity',
                        'Eigenkapital',
                        'GesamtEigenkapital'
                    ]
                }
            },
            'total_capital': {
                'statements': ['BS'],
                'taxonomies': {
                    'ifrs-full': [
                        'TotalCapital',
                        'TotalCapitalisation',
                        'CapitalAndReserves',
                        'EquityAndBorrowings'
                    ],
                    'us-gaap': [
                        'TotalCapital',
                        'TotalCapitalization',
                        'StockholdersEquityIncludingPortionAttributableToNoncontrollingInterest',
                    ],
                    'generic': [
                        'TotalCapital',
                        'Gesamtkapital',
                        'KapitalGesamt',
                        'GesamtKapitalisierung'
                    ]
                }
            },
            'financial_liabilities': {
                'statements': ['BS'],
                'taxonomies': {
                    'ifrs-full': [
                        'FinancialLiabilities',
                        'FinancialLiabilitiesAtAmortisedCost',
                        'FinancialLiabilitiesAtFairValue',
                        'NoncurrentFinancialLiabilities',
                        'CurrentFinancialLiabilities'
                    ],
                    'us-gaap': [
                        'FinancialLiabilities',
                        'LongTermDebtAndCapitalLeaseObligations',
                        'ShortTermDebtAndCapitalLeaseObligations'
                    ],
                    'generic': [
                        'TotalFinancialLiabilities',
                        'FinanzielleVerbindlichkeiten',
                        'FinanzschuldenGesamt'
                    ]
                }
            },
            'bonds_payable': {
                'statements': ['BS'],
                'taxonomies': {
                    'ifrs-full': [
                        'BondsIssued',
                        'DebtInstrumentsIssued',
                        'NoncurrentPortionOfNoncurrentBondsIssued',
                        'CurrentBondsIssuedAndCurrentPortionOfNoncurrentBondsIssued',
                    ],
                    'us-gaap': [
                        'ConvertibleDebtCurrent',
                        'ConvertibleDebtNoncurrent',
                        'SeniorNotes'
                    ],
                    'generic': [
                        'Anleihen',
                        'AnleihenVerbindlichkeiten',
                        'SchuldverschreibungenBegeben'
                    ]
                }
            },
            'lease_liabilities': {
                'statements': ['BS'],
                'taxonomies': {
                    'ifrs-full': [
                        'LeaseLiabilities',
                        'CurrentLeaseLiabilities',
                        'NoncurrentLeaseLiabilities'
                    ],
                    'us-gaap': [
                        'OperatingLeaseLiability',
                        'FinanceLeaseObligations',
                        'LeaseLiability'
                    ],
                    'generic': [
                        'Leasingverbindlichkeiten',
                        'LeasingVerpflichtungen'
                    ]
                }
            },
            'bank_loans': {
                'statements': ['BS'],
                'taxonomies': {
                    'ifrs-full': [
                        'BankBorrowings',
                        'LoansAndAdvances',
                        'BankLoansPayable'
                    ],
                    'us-gaap': [
                        'BankDebt',
                        'LineOfCredit',
                        'NotesPayableToBanks'
                    ],
                    'generic': [
                        'Loans',
                        'Bankkredite',
                        'Bankdarlehen',
                        'KrediteBanken'
                    ]
                }
            },   
            'long_term_borrowings': {
                'statements': ['BS'],
                'taxonomies': {
                    'ifrs-full': [
                        'NoncurrentBorrowings',
                        'LongTermBorrowings',
                        'LongtermBorrowings',
                        'NoncurrentInterestBearingBorrowings'
                        'TotalDebt',
                        'NonCurrentBorrowings',                 
                    ],
                    'us-gaap': [
                        'LongTermDebtNoncurrent',
                        'LongTermDebtAndCapitalLeaseObligations'
                    ],
                    'generic': [
                        'LangfristigeVerbindlichkeiten',
                        'LangfristigeFinanzschulden'
                    ]
                }
            },
            'short_term_borrowings': {
                'statements': ['BS'],
                'taxonomies': {
                    'ifrs-full': [
                        'CurrentBorrowings',
                        'ShortTermBorrowings',
                        'ShorttermBorrowings',
                        'CurrentInterestBearingBorrowings',  
                        'CurrentBorrowingsAndCurrentPortionOfNoncurrentBorrowings',
                        'CurrentPortionOfLongtermBorrowings',
                    ],
                    'us-gaap': [
                        'ShortTermBorrowings',
                        'LongTermDebtCurrent'
                    ],
                    'generic': [
                        'KurzfristigeVerbindlichkeiten',
                        'KurzfristigeFinanzschulden'
                    ]
                }
            },
            'share_capital': {
                'statements': ['BS'],
                'taxonomies': {
                    'ifrs-full': [
                        'IssuedCapital',
                        'ShareCapital'
                    ],
                    'us-gaap': [
                        'CommonStockValue',
                        'PreferredStockValue'
                    ],
                    'generic': [
                        'Grundkapital',
                        'Stammaktien'
                    ]
                }
            },
            'retained_earnings': {
                'statements': ['BS'],
                'taxonomies': {
                    'ifrs-full': [
                        'RetainedEarnings',
                        'AccumulatedRetainedEarnings'
                    ],
                    'us-gaap': [
                        'RetainedEarningsAccumulatedDeficit',
                        'AccumulatedDeficit'
                    ],
                    'generic': [
                        'Gewinnrücklagen',
                        'BilanzverlustGewinn'
                    ]
                }
            },
            'reserves': {
                'statements': ['BS'],
                'taxonomies': {
                    'ifrs-full': [
                        'OtherReserves',
                        'Reserves'
                    ],
                    'us-gaap': [
                        'AccumulatedOtherComprehensiveIncomeLoss',
                        'AdditionalPaidInCapital'
                    ],
                    'generic': [
                        'Kapitalrücklage',
                        'SonstigeRücklagen'
                    ]
                }
            } , 
            'cash_and_equivalents': {
                'statements': ['BS'],
                'taxonomies': {
                    'ifrs-full': [
                        'CashAndCashEquivalents',
                        'RestrictedCashAndCashEquivalents',
                        'CashAndBankBalancesAtCentralBanks',
                        'CashOnHand',
                        'BalancesWithBanks'
                    ],
                    'us-gaap': [
                        'CashAndCashEquivalentsAtCarryingValue',
                        'Cash',
                        'CashEquivalents',
                        'RestrictedCashAndCashEquivalents',
                        'CashAndDueFromBanks'
                    ],
                    'generic': [
                        'TotalCashAndEquivalents',
                        'CashAndBankBalances',
                        'LiquidFunds',
                        'ZahlungsmittelUndZahlungsmitteläquivalente',  # German
                        'FlüssigeMittel',                              # German
                        'BankguthabenUndKassenbestand',                # German
                        'LiquideMittel',                               # German
                        'KasseUndBankguthaben'                         # German
                    ]
                }
            },            
            'cash': {
                'statements': ['BS'],
                'taxonomies': {
                    'ifrs-full': [
                        'CashOnHand',
                        'Cash',
                        'CashInBank',
                        'CashBalancesAtCentralBanks'
                    ],
                    'us-gaap': [
                        'Cash',
                        'CashOnHand',
                        'CashInBanks',
                        'UnrestrictedCash'
                    ],
                    'generic': [
                        'Kassenbestand',
                        'Bankguthaben',
                        'BarMittel',
                        'KasseGuthaben',
                        'Kassenmittel'
                    ]
                }
            },
            'cash_equivalents': {
                'statements': ['BS'],
                'taxonomies': {
                    'ifrs-full': [
                        'CashEquivalents',
                        'ShortTermInvestments',
                        'ShortTermDeposits',
                        'OtherCashAndCashEquivalents'
                    ],
                    'us-gaap': [
                        'CashEquivalents',
                        'CashEquivalentsAtCarryingValue',
                        'ShortTermInvestments',
                        'MarketableSecurities'
                    ],
                    'generic': [
                        'Zahlungsmitteläquivalente',
                        'KurzfristigeGeldanlagen',
                        'GeldmarktAnlagen',
                        'LiquideWertpapiere',
                        'KurzfristigeFinanzanlagen'
                    ]
                }
            },
            'trade_receivables': {
                'statements': ['BS'],
                'taxonomies': {
                    'ifrs-full': [
                        'TradeAndOtherCurrentReceivables',
                        'TradeReceivables'
                    ],
                    'us-gaap': [
                        'AccountsReceivableNetCurrent',
                        'AccountsNotesAndLoansReceivableNetCurrent'
                    ],
                    'generic': [
                        'Forderungen',
                        'ForderungenAusLieferungenUndLeistungen'
                    ]
                }
            },
            'marketable_securities': {
                'statements': ['BS'],
                'taxonomies': {
                    'ifrs-full': [
                        'CurrentFinancialAssets',
                        'CurrentFinancialAssetsAtFairValueThroughProfitOrLoss',
                        'OtherCurrentFinancialAssets'
                    ],
                    'us-gaap': [
                        'MarketableSecuritiesCurrent',
                        'AvailableForSaleSecuritiesCurrent',
                        'TradingSecuritiesCurrent'
                    ],
                    'generic': [
                        'Wertpapiere',
                        'KurzfristigeWertpapiere',
                        'HandelbarePapiere'
                    ]
                }
            },
            'quick_assets': {
                'statements': ['BS'],
                'taxonomies': {
                    'ifrs-full': [
                        'QuickAssets', #dummy
                    ],
                    'us-gaap': [
                        'QuickAssets', #dummy
                    ],
                    'generic': [
                        'SchnellLiquideVermögenswerte',
                        'LiquideMittelUndForderungen'
                    ]
                }
            },
            
            'operating_expenses': {
                'statements': ['IS'],
                'taxonomies': {
                    'ifrs-full': [
                        'OperatingExpense',
                        'SellingGeneralAndAdministrativeExpense',
                        'AdministrativeExpense',
                        'DistributionCosts'
                    ],
                    'us-gaap': [
                        'OperatingExpenses',
                        'SellingGeneralAndAdministrativeExpense',
                        'GeneralAndAdministrativeExpense'
                    ],
                    'generic': [
                        'Betriebsaufwand',
                        'Verwaltungsaufwand',
                        'Vertriebskosten'
                    ]
                }
            },
            'depreciation_amortization': {
                'statements': ['IS'],
                'taxonomies': {
                    'ifrs-full': [
                        'DepreciationAndAmortisationExpense',
                        'DepreciationExpense',
                        'AmortisationExpense'
                    ],
                    'us-gaap': [
                        'DepreciationAndAmortization',
                        'DepreciationDepletionAndAmortization'
                    ],
                    'generic': [
                        'AbschreibungenAmortisationen',
                        'Abschreibungen',
                        'AbschreibungenWertminderungen'
                    ]
                }
            },
            'interest_expense': {
                'statements': ['IS'],
                'taxonomies': {
                    'ifrs-full': [
                        'InterestExpense',
                        'FinanceCosts',
                        'FinanceExpense'
                    ],
                    'us-gaap': [
                        'InterestExpense',
                        'InterestPaid',
                        'InterestCostsIncurred'
                    ],
                    'generic': [
                        'Zinsaufwand',
                        'Finanzaufwand',
                        'Finanzierungskosten'
                    ]
                }
            },
            'income_tax_expense': {
                'statements': ['IS'],
                'taxonomies': {
                    'ifrs-full': [
                        'IncomeTaxExpenseContinuingOperations',
                        'TaxExpense',
                        'CurrentTaxExpense'
                    ],
                    'us-gaap': [
                        'IncomeTaxExpenseBenefit',
                        'CurrentIncomeTaxExpense'
                    ],
                    'generic': [
                        'Steueraufwand',
                        'Ertragsteuern',
                        'Einkommenssteuer'
                    ]
                }
            },
            'ebit': {
                'statements': ['IS'],
                'taxonomies': {
                    'ifrs-full': [
                        'EBIT', #dummy
                    ],
                    'us-gaap': [
                        'EBIT', #dummy
                    ],
                    'generic': [
                        'EBIT', 
                    ]
                }
            },
            'ebt': {
                'statements': ['IS'],
                'taxonomies': {
                    'ifrs-full': [
                        'EBT', #dummy
                    ],
                    'us-gaap': [
                        'EBT', #dummy
                    ],
                    'generic': [
                        'EBT',
                    ]
                }
            },
            'ebitda': {
                'statements': ['IS'],
                'taxonomies': {
                    'ifrs-full': [
                        'EBITDA', #dummy
                    ],
                    'us-gaap': [
                        'EBITDA', #dummy
                    ],
                    'generic': [
                        'EBITDA',
                    ]
                }
            },      
            
            'operating_profit': {
                'statements': ['IS'],
                'taxonomies': {
                    'ifrs-full': [
                        'ProfitLossFromOperatingActivities',
                        'OperatingIncome',
                        'OperatingProfitLoss'
                    ],
                    'us-gaap': [
                        'OperatingIncomeLoss',
                        'IncomeLossFromContinuingOperations',
                        'IncomeLossFromOperations'
                    ],
                    'generic': [
                        'OperatingProfit',
                        'OperatingIncome',
                        'BetriebsErgebnis',
                        'OperativesErgebnis'
                    ]
                }
            },

            'capital_employed': {
                'statements': ['BS'],
                'taxonomies': {
                    'ifrs-full': [
                        'CapitalEmployed',
                        'NetOperatingAssets'
                    ],
                    'us-gaap': [
                        'NetOperatingAssets',
                        'CapitalEmployed'
                    ],
                    'generic': [
                        'TotalCapitalEmployed',
                        'EingesetztesKapital',
                        'BetriebsnotwendigesKapital'
                    ]
                }
            },
              
            'inventory': {
                'statements': ['BS'],
                'taxonomies': {
                    'ifrs-full': [
                        'Inventories',
                        'CurrentInventories',
                        'InventoriesTotal'
                    ],
                    'us-gaap': [
                        'InventoryNet',
                        'Inventories',
                        'InventoriesAndContractsInProgress'
                    ],
                    'generic': [
                        'Vorräte',
                        'VorräteGesamt',
                        'Lagerbestand'
                    ]
                }
            },
            'dividends': {
                'statements': ['CF', 'IS'],
                'taxonomies': {
                    'ifrs-full': [
                        'DividendsPaid',
                        'DividendsPaidClassifiedAsFinancingActivities',
                        'DividendsRecognisedAsDistributionsToOwnersOfParent'
                    ],
                    'us-gaap': [
                        'PaymentsOfDividends',
                        'DividendsPaidCommonStock',
                        'DividendsDeclared'
                    ],
                    'generic': [
                        'Dividenden',
                        'DividendenZahlungen',
                        'AusschüttungenEigentümer'
                    ]
                }
            },
            'net_debt': {
                'statements': ['BS'],
                'taxonomies': {
                    'ifrs-full': [
                    ],
                    'us-gaap': [
                    ],
                    'generic': [
                    ]
                }
            }                                                                      
            #
                                                
            # Add more concepts as needed...
        }


class ValueMetadata:
    """Metadata for tracked financial values"""
    def __init__(self, qname: str = None, source_statement: str = None, 
                 is_custom: bool = False, components: Dict[str, float] = None, 
                 calculation_rule: str = None, last_updated: datetime = None,
                 context_id: str = None):
        self.qname = qname
        self.source_statement = source_statement
        self.is_custom = is_custom
        self.components = components or {}
        self.calculation_rule = calculation_rule
        self.last_updated = last_updated or datetime.now()
        self.context_id = context_id
    
    def to_dict(self) -> dict:
        return {
            'qname': self.qname,
            'source': self.source_statement,
            'is_custom': self.is_custom,
            'components': self.components,
            'calculation_rule': self.calculation_rule,
            'last_updated': self.last_updated.isoformat(),
            'context_id': self.context_id
        }
        



class ComponentRegistry:
    """
    Registry of compound financial value definitions with dependency tracking and validation.
    
    Manages:
    - Component hierarchies
    - Calculation rules
    - Dependency relationships
    - Validation rules
    """
    
    def __init__(self):
        self._registry = {}
        self._dependencies = defaultdict(set)
        self._validation_rules = {}
        self.logger = logger
        
        log_debug("\nInitializing component registry...")
        self._initialize_standard_components()
        
    def _initialize_standard_components(self):
        """Initialize registry with standard financial compound values"""
        # Balance Sheet components
        self.register_compound('total_debt', 
            ['long_term_borrowings', 'short_term_borrowings', 'lease_liabilities', 
             'bonds_payable', 'bank_loans'],
            calculation_rule="sum",     
            validation_rules={
                'non_negative': True,
                'required_components': ['long_term_borrowings', 'short_term_borrowings']
            })
            
        self.register_compound('total_capital', 
            ['total_debt', 'total_equity'],
            validation_rules={
                'non_negative': True,
                'required_components': ['total_equity']
            })
            
        self.register_compound('working_capital', 
            ['current_assets', 'current_liabilities'],
            calculation_rule="subtract",
            validation_rules={
                'required_components': ['current_assets', 'current_liabilities']
            })
        
        self.register_compound('quick_assets', 
            ['cash_and_equivalents', 'trade_receivables', 'marketable_securities'],
            calculation_rule="sum",            
            validation_rules={
                'non_negative': True,
                'required_components': ['cash_and_equivalents', 'trade_receivables']
            })        
            
        # Income Statement components
        self.register_compound('gross_profit', 
            ['revenue', 'cost_of_sales'],
            calculation_rule="subtract",
            validation_rules={
                'required_components': ['revenue', 'cost_of_sales']
            })

        self.register_compound('cash_and_equivalents',
            ['cash', 'cash_equivalents'],
            validation_rules={
                'non_negative': True
            })

        self.register_compound('return_on_equity',
            ['net_income', 'total_equity'],
            calculation_rule="divide",
            validation_rules={
                'required_components': ['net_income', 'total_equity']
            })            

        self.register_compound('ebit',
            ['revenue', 'cost_of_sales', 'operating_expenses'],
            calculation_rule="subtract",
            validation_rules={
                'required_components': ['revenue', 'cost_of_sales']
            })

        self.register_compound('ebitda',
            ['ebit', 'depreciation_amortization'],
            calculation_rule="sum",
            validation_rules={
                'required_components': ['ebit', 'depreciation_amortization'],
                'non_zero_components': ['depreciation_amortization']  # Only calculate if DA exists
            })

        self.register_compound('ebt',
            ['ebit', 'interest_expense'],
            calculation_rule="subtract",
            validation_rules={
                'required_components': ['ebit','interest_expense']
            })

        self.register_compound('net_debt',
            ['total_debt', 'cash_and_equivalents'],
            calculation_rule="subtract",
            validation_rules={
                'required_components': ['total_debt', 'cash_and_equivalents'],
                'non_zero_components': ['total_debt', 'cash_and_equivalents']  # Only calculate if exists
            })
                
        log_debug(f"Registered compounds: {list(self._registry.keys())}")
        
    def register_compound(self, name: str, components: List[str], 
                         calculation_rule: str = "sum",
                         validation_rules: dict = None):
        """
        Register a compound value definition with enhanced validation
        
        Args:
            name: Name of compound value (e.g., 'total_debt')
            components: List of component names
            calculation_rule: How to combine components (sum/subtract)
            validation_rules: Dictionary of validation rules
        """
        if name in self._registry:
            self.logger.log(f"Updating existing compound: {name}")
            
        self._registry[name] = {
            'components': components,
            'calculation_rule': calculation_rule,
            'validation_rules': validation_rules or {}
        }
        
        # Update dependency graph
        for component in components:
            self._dependencies[component].add(name)
            
        self.logger.log(f"Registered compound {name} with {len(components)} components")
        
    def get_components(self, name: str) -> List[str]:
        """Get list of components for a compound value"""
        return self._registry.get(name, {}).get('components', [])
        
    def get_calculation_rule(self, name: str) -> str:
        """Get calculation rule for a compound value"""
        return self._registry.get(name, {}).get('calculation_rule', 'sum')
        
    def get_validation_rules(self, name: str) -> dict:
        """Get validation rules for a compound value"""
        return self._registry.get(name, {}).get('validation_rules', {})
        
    def get_dependents(self, name: str) -> Set[str]:
        """Get all values that depend on this value"""
        return self._dependencies[name]
        
    def validate_compound_value(self, name: str, values: dict) -> Tuple[bool, List[str]]:
        """
        Validate a compound value calculation against its rules
        
        Args:
            name: Compound value name
            values: Dictionary of component values
            
        Returns:
            Tuple of (is_valid, list of validation messages)
        """
        if name not in self._registry:
            return False, [f"Unknown compound value: {name}"]
            
        rules = self.get_validation_rules(name)
        messages = []
        
        # Check required components
        if 'required_components' in rules:
            for required in rules['required_components']:
                if required not in values:
                    messages.append(f"Missing required component: {required}")
                    
        # Check non-negative rule
        if rules.get('non_negative', False):
            for component, value in values.items():
                if value is not None and value < 0:
                    messages.append(f"Negative value not allowed: {component}={value}")
                    
        return len(messages) == 0, messages
        
    def check_circular_dependency(self, name: str, chain: Set[str] = None) -> bool:
        """Check for circular dependencies in compound value definition"""
        if chain is None:
            chain = set()
            
        if name in chain:
            return True
            
        chain.add(name)
        
        for component in self.get_components(name):
            if component in self._registry:
                if self.check_circular_dependency(component, chain):
                    return True
                    
        chain.remove(name)
        return False
        
    def get_compound_info(self, name: str) -> dict:
        """Get comprehensive information about a compound value"""
        if name not in self._registry:
            return None
            
        info = self._registry[name].copy()
        info['dependents'] = list(self.get_dependents(name))
        info['has_circular_dependency'] = self.check_circular_dependency(name)
        
        return info
        
    def list_all_compounds(self) -> List[dict]:
        """Get information about all registered compounds"""
        return [
            {
                'name': name,
                'info': self.get_compound_info(name)
            }
            for name in self._registry.keys()
        ]
        
class ValueCache:
    """
    Enhanced cache for financial values with metadata tracking and validation.
    
    Features:
    - Period-aware caching
    - Metadata preservation
    - Cache statistics
    - Invalidation tracking
    - TTL support
    """
    
    def __init__(self, ttl_seconds: int = 3600):
        self._values = {}  # (key, period) -> value
        self._metadata = {}  # (key, period) -> metadata
        self._context_index = {}  # (key, context_id) -> period
        self._access_times = {}
        self.ttl_seconds = ttl_seconds
        self.logger = logger
        log_debug("\nInitialized ValueCache with context indexing")
        
    def get(self, key: str, period: tuple) -> Tuple[Optional[float], Optional[ValueMetadata]]:
        """
        Get cached value and its metadata with TTL check
        
        Args:
            key: Concept identifier
            period: (start_date, end_date) tuple
            
        Returns:
            Tuple of (value, metadata) or (None, None) if not found/expired
        """
        cache_key = (key, period)
        
        # Check if value exists and isn't expired
        if cache_key in self._values:
            last_access = self._access_times.get(cache_key)
            if last_access and (datetime.now() - last_access).total_seconds() < self.ttl_seconds:
                self._access_times[cache_key] = datetime.now()  # Update access time
                return self._values.get(cache_key), self._metadata.get(cache_key)
            else:
                # Expired - remove it
                self.invalidate(key, period, reason="TTL expired")
                
        return None, None
        
    def set(self, concept: str, period: tuple, value: float, metadata: ValueMetadata):
        """Set value with context tracking"""
        cache_key = (concept, period)
        self._values[cache_key] = value
        self._metadata[cache_key] = metadata
        self._access_times[cache_key] = datetime.now()
        
        # Add context index if available
        if metadata and metadata.context_id:
            self._context_index[(concept, metadata.context_id)] = period
            # log_debug(f"Cached mapping: {metadata.context_id} -> {period}")
        
    def invalidate(self, key: str, period: tuple, reason: str = None):
        """
        Invalidate a cached value and track the invalidation
        
        Args:
            key: Concept identifier
            period: (start_date, end_date) tuple
            reason: Optional reason for invalidation
        """
        cache_key = (key, period)
        
        # Track invalidation
        if cache_key in self._values:
            self._invalidation_log.append({
                'key': key,
                'period': period,
                'value': self._values[cache_key],
                'metadata': self._metadata.get(cache_key),
                'reason': reason,
                'timestamp': datetime.now()
            })
            
            # Remove from all tracking
            self._values.pop(cache_key, None)
            self._metadata.pop(cache_key, None)
            self._access_times.pop(cache_key, None)
            
    def invalidate_dependents(self, key: str, component_registry):
        """
        Invalidate all values that depend on the given key
        
        Args:
            key: Concept identifier
            component_registry: ComponentRegistry instance for dependency checking
        """
        dependents = component_registry.get_dependents(key)
        for dependent in dependents:
            # Invalidate dependent for all periods
            dependent_keys = [k for k in self._values.keys() if k[0] == dependent]
            for dep_key in dependent_keys:
                self.invalidate(dep_key[0], dep_key[1], 
                              reason=f"Dependent of invalidated {key}")
                
    def get_period_values(self, key: str) -> Dict[tuple, Tuple[float, ValueMetadata]]:
        """
        Get all values for a concept across different periods
        
        Args:
            key: Concept identifier
            
        Returns:
            Dictionary mapping periods to (value, metadata) tuples
        """
        return {
            period: (self._values.get((key, period)),
                    self._metadata.get((key, period)))
            for period in self.get_periods(key)
        }
        
    def get_periods(self, key: str) -> Set[tuple]:
        """Get all periods available for a concept"""
        return {period for k, period in self._values.keys() if k == key}
        
    def get_cache_stats(self) -> dict:
        """Get comprehensive cache statistics"""
        current_time = datetime.now()
        return {
            'total_entries': len(self._values),
            'total_metadata': len(self._metadata),
            'unique_concepts': len({k[0] for k in self._values.keys()}),
            'unique_periods': len({k[1] for k in self._values.keys()}),
            'invalidation_count': len(self._invalidation_log),
            'expired_count': sum(
                1 for t in self._access_times.values()
                if (current_time - t).total_seconds() >= self.ttl_seconds
            )
        }
        
    def clear_expired(self):
        """Remove all expired entries"""
        current_time = datetime.now()
        expired_keys = [
            k for k, t in self._access_times.items()
            if (current_time - t).total_seconds() >= self.ttl_seconds
        ]
        
        for key in expired_keys:
            self.invalidate(key[0], key[1], reason="Expired in cleanup")
            
    def clear(self):
        """Clear all cached values"""
        self._values.clear()
        self._metadata.clear()
        self._access_times.clear()
        
    def get_invalidation_history(self) -> List[dict]:
        """Get the invalidation history"""
        return sorted(
            self._invalidation_log,
            key=lambda x: x['timestamp'],
            reverse=True
        )

    def get_by_context(self, concept: str, context_id: str) -> Tuple[Optional[float], Optional[ValueMetadata]]:
        """Get value using context ID"""
        # log_debug(f"\nAttempting to get value for {concept} using context {context_id}")
        
        # Check if we have this context mapping
        if (concept, context_id) in self._context_index:
            period = self._context_index[(concept, context_id)]
            # log_debug(f"Found period mapping for context: {period}")
            return self.get(concept, period)
            
        # log_debug(f"No context mapping found for {context_id}")
        return None, None

class FinancialDataManager:
    """
    Central manager for financial data access and compound value calculations.
    Provides unified interface for value retrieval, caching, and tracking.
    """
    
    def __init__(self, fact_provider):
        self.fact_provider = fact_provider
        self.component_registry = ComponentRegistry()
        self.value_cache = ValueCache()
        self.context_mappings = {}
        self.calculation_tracker = CalculationTracker()
        self.access_tracker = AccessTracker()  
        self.calculation_inputs = defaultdict(list)  # Add this line
        self.composition_registry = defaultdict(dict)  # Add this line if not present
        self.logger = logger
        log_debug("\nInitialized FinancialDataManager with context mappings (FinancialDataManager.__init__)")

    def initialize_contexts(self):
        """Initialize contexts after table facts have been assigned"""
        log_debug("\n=== Initializing Statement Contexts (FilingAnalyzer.initialize_contexts) ===")
        self.statement_periods = self._initialize_statement_contexts()
        populated_statements = {stmt: len(contexts) 
                            for stmt, contexts in self.statement_periods.items() 
                            if contexts}
        log_debug("\nPopulated statement contexts:")
        for stmt, count in populated_statements.items():
            log_debug(f"- {stmt}: {count} contexts")
    
    def get_value(self, concept: str, period: tuple, statement_type: str) -> Tuple[Optional[float], ValueMetadata]:
        """Primary method for getting financial values"""
        # log_debug(f"\nRetrieving value for {concept} in period {period}")
        try:
            # Track initial access attempt
            self.access_tracker.track_access(concept, 'managed', {
                'period': period,
                'statement_type': statement_type,
                'method': 'get_value'
            })

            # Check cache first
            cached_value, cached_metadata = self.value_cache.get(concept, period)
            if cached_value is not None:
                self.access_tracker.track_access(concept, 'cached', {
                    'period': period,
                    'metadata': cached_metadata.to_dict() if cached_metadata else None
                })
                return cached_value, cached_metadata

            # Validate access
            if not self._validate_access_request(concept, period, statement_type):
                log_debug(f"Invalid access request for {concept}")
                return None, None

            # Find matching context for period
            context_id = self._get_context_for_period(period, statement_type)
            if not context_id:
                log_debug(f"No matching context found for period {period}")
                return None, None

            # Try direct lookup first
            value, qname, is_custom = self.fact_provider.get_fact_value_by_context(
                concept, context_id, statement_type)
                
            if value is not None:
                metadata = ValueMetadata(
                    qname=qname,
                    source_statement=statement_type,
                    is_custom=is_custom,
                    context_id=context_id
                )
                self.value_cache.set(concept, period, value, metadata)
                return value, metadata

            # If direct lookup fails and it's a compound value, try calculation
            if concept in self.component_registry._registry:
                return self._get_compound_value(concept, period, statement_type)

            return None, None

        except Exception as e:
            self.logger.log(f"Error getting value for {concept}: {str(e)}", include_trace=True)
            return None, None

    def _get_compound_value(self, concept: str, period: tuple, statement_type: str, dependency_chain: Set[str] = None) -> Tuple[Optional[float], ValueMetadata]:
        """Calculate a compound value from its components with comprehensive tracking"""
        log_debug(f"\n=== Getting Compound Value for {concept} ===")
        log_debug(f"Calculation rule: {self.component_registry.get_calculation_rule(concept)}")
        log_debug(f"Components: {self.component_registry.get_components(concept)}")
        
        if dependency_chain is None:
            dependency_chain = set()
                
        if concept in dependency_chain:
            self.logger.log(f"Circular dependency detected for {concept}", "ERROR")
            return None, None
                
        dependency_chain.add(concept)
        
        try:
            components = self.component_registry.get_components(concept)
            calculation_rule = self.component_registry.get_calculation_rule(concept)
            validation_rules = self.component_registry.get_validation_rules(concept)
            
            # Verify required components with values
            if validation_rules.get('required_components'):
                missing_components = []
                for req_component in validation_rules['required_components']:
                    value, metadata = self.get_value(req_component, period, statement_type)
                    if value is None:
                        missing_components.append(req_component)
                        
                if missing_components:
                    log_debug(f"Missing required components for {concept}: {missing_components}")
                    return None, None
                    
            # Check non-zero components
            if validation_rules.get('non_zero_components'):
                zero_components = []
                for comp in validation_rules['non_zero_components']:
                    value, metadata = self.get_value(comp, period, statement_type)
                    if value == 0:
                        zero_components.append(comp)
                        
                if zero_components:
                    log_debug(f"Zero value components for {concept}: {zero_components}")
                    return None, None     
                       
            component_values = {}
            component_metadata = {}
            
            # Get all component values
            for i, component in enumerate(components):
                value, metadata = self.get_value(component, period, statement_type)
                if value is not None:
                    component_values[component] = value
                    component_metadata[component] = metadata
                    
                    if hasattr(self.fact_provider, 'ratio_calculator'):
                        # Determine sign based on calculation rule
                        sign = "+"  # Default for first component and sum rule
                        if calculation_rule == "subtract" and i > 0:
                            sign = "-"
                        elif calculation_rule == "divide":
                            sign = "/" if i == 1 else "+"  # Division symbol for divisor
                        
                        self.fact_provider.ratio_calculator.input_tracker.add_input(
                            name=component,
                            qname=metadata.qname if metadata else 'Unknown',
                            value=value,
                            context_id=metadata.context_id if metadata else None,
                            statement=statement_type,
                            indent_level=1
                            # sign=sign
                        )
            
            # Calculate if we have components
            if component_values:
                total_value = None
                if calculation_rule == "sum":
                    total_value = sum(component_values.values())
                elif calculation_rule == "subtract":
                    if len(component_values) >= 2:
                        values = list(component_values.values())
                        total_value = values[0] - sum(values[1:])
                elif calculation_rule == "divide":
                    if len(component_values) >= 2:
                        values = list(component_values.values())
                        if values[1] != 0:  # Check for division by zero
                            total_value = values[0] / values[1]
                        else:
                            log_debug(f"Division by zero prevented for {concept}")
                        
                if total_value is not None:
                    metadata = ValueMetadata(
                        qname="Calculated",
                        source_statement=statement_type,
                        is_custom=all(m.is_custom for m in component_metadata.values()),
                        components=component_values,
                        calculation_rule=calculation_rule
                    )
                    
                    # Track the compound value itself
                    if hasattr(self.fact_provider, 'ratio_calculator'):
                        self.fact_provider.ratio_calculator.input_tracker.add_input(
                            name=concept,
                            qname=metadata.qname,
                            value=total_value,
                            context_id=next(iter(component_metadata.values())).context_id if component_metadata else None,
                            statement=statement_type,
                            indent_level=0
                        )
                    
                    return total_value, metadata
                
            return None, None
                
        finally:
            dependency_chain.remove(concept)            
            
    def record_value_usage(self, concept: str, period: tuple, value: float, 
                          metadata: ValueMetadata, is_component: bool = False):
        """Record value usage for tracking and audit"""
        self.calculation_inputs[period].append({
            'concept': concept,
            'value': value,
            'metadata': metadata,
            'is_component': is_component,
            'timestamp': datetime.now()
        })

    def record_composition(self, compound: str, period: tuple, total_value: float, 
                         components: dict):
        """Record composition of compound values for transparency"""
        self.composition_registry[period][compound] = {
            'value': total_value,
            'components': components,
            'calculated_at': datetime.now()
        }

    def get_value_composition(self, concept: str, period: tuple) -> dict:
        """Get detailed composition of a value including all components"""
        value, metadata = self.value_cache.get(concept, period)
        
        if value is None:
            value, metadata = self.get_value(concept, period, 
                'BS' if concept in ['total_debt', 'total_capital', 'working_capital','quick_assets','net_debt'] else 'IS')
            
        if value is None:
            return None
            
        composition = {
            'value': value,
            'metadata': metadata.to_dict() if metadata else None,
            'components': {}
        }
        
        if metadata and metadata.components:
            for component, component_value in metadata.components.items():
                component_composition = self.get_value_composition(component, period)
                if component_composition:
                    composition['components'][component] = component_composition
        
        return composition

    def validate_matching_periods(self, period_bs: tuple, period_is: tuple) -> bool:
        """Validate that BS and IS periods are compatible for ratio calculation"""
        if not period_bs[1] or not period_is[1]:
            return False
            
        if period_bs[0] is not None:
            return False
            
        if period_is[0] is None:
            return False
            
        return period_bs[1] == period_is[1]
     
    def get_calculation_details(self) -> dict:
        """Get comprehensive calculation information for debugging and audit"""
        return {
            'inputs': dict(self.calculation_inputs),
            'compositions': dict(self.composition_registry),
            'cache_stats': {
                'size': len(self.value_cache._values),
                'metadata_count': len(self.value_cache._metadata)
            }
        }

    def get_period_key(self, context) -> Optional[tuple]:
        """
        Get standardized period key from a context.
        
        Args:
            context: XBRL context object
            
        Returns:
            Tuple of (start_date, end_date) or (None, end_date) for instants
        """
        if not context:
            return None
            
        if context.isInstantPeriod and context.instantDatetime:
            return (None, context.instantDatetime)
        elif context.isStartEndPeriod and context.startDatetime and context.endDatetime:
            return (context.startDatetime, context.endDatetime)
            
        return None

    def get_period_dates(self, context) -> Tuple[Optional[datetime], Optional[datetime]]:
        """
        Extract start and end dates from context.
        
        Args:
            context: XBRL context object
            
        Returns:
            Tuple of (start_date, end_date)
        """
        if not context:
            return None, None
            
        start_date = context.startDatetime if context.isStartEndPeriod else None
        end_date = context.endDatetime if context.endDatetime else context.instantDatetime
            
        return start_date, end_date
    
    def _get_context_for_period(self, period: tuple, statement_type: str) -> Optional[str]:
        """Find matching context for a given period"""
        # log_debug(f"\n=== Finding context for period {period} (type: {statement_type}) ===")
        contexts = self.fact_provider.filing_analyzer.statement_periods.get(statement_type, {})
        # log_debug(f"Available contexts for {statement_type}: {list(contexts.keys())}")
        
        for context_id, ctx_period in contexts.items():
            # log_debug(f"Checking context {context_id}: {ctx_period}")
            if ctx_period == period:
                # log_debug(f"Found matching context: {context_id}")
                return context_id
                
        log_debug("No matching context found")
        return None

    def _validate_access_request(self, concept: str, period: tuple, statement_type: str) -> bool:
        """Validate value access request"""
        try:
            # Validate concept
            if concept not in self.component_registry._registry and not self._is_valid_concept(concept):
                log_debug(f"Invalid concept: {concept}")
                return False

            # Validate period
            if not self._is_valid_period(period, statement_type):
                log_debug(f"Invalid period: {period} for {statement_type}")
                return False

            # Validate statement type
            if not statement_type in ['BS', 'IS', 'CF', 'CI', 'EQ']:
                log_debug(f"Invalid statement type: {statement_type}")
                return False

            return True

        except Exception as e:
            self.logger.log(f"Error in access validation: {str(e)}", include_trace=True)
            return False

    def _is_valid_concept(self, concept: str) -> bool:
        """Check if concept is valid"""
        return concept in self.fact_provider.concept_maps

    def _is_valid_period(self, period: tuple, statement_type: str) -> bool:
        """Check if period is valid for statement type"""
        if not period or len(period) != 2:
            return False

        # Get valid periods for statement type
        valid_periods = set()
        for context_id, ctx_period in self.fact_provider.filing_analyzer.statement_periods.get(statement_type, {}).items():
            valid_periods.add(ctx_period)

        return period in valid_periods

    def get_value_by_context(self, concept: str, context_id: str, statement_type: str) -> Tuple[Optional[float], Optional[ValueMetadata]]:
        """Get value using context ID with caching and tracking"""
        # Try cache first
        value, metadata = self.value_cache.get_by_context(concept, context_id)
        if value is not None:
            return value, metadata

        # Get period
        period = self.fact_provider.filing_analyzer.statement_periods[statement_type].get(context_id)
        if not period:
            return None, None

        # Try direct lookup first
        direct_value, qname, is_custom = self.fact_provider.get_fact_value_by_context(concept, context_id, statement_type)
        if direct_value is not None:
            metadata = ValueMetadata(qname=qname, source_statement=statement_type, is_custom=is_custom, context_id=context_id)
            self.value_cache.set(concept, period, direct_value, metadata)
            return direct_value, metadata

        # Only try compound calculation if direct lookup failed
        if concept in self.component_registry._registry:
            return self._get_compound_value(concept, period, statement_type)

        return None, None

    def get_context_matches(self, context_id: str, target_statement: str) -> List[str]:
        """Get matching contexts for a given context ID"""
        if context_id in self.context_mappings:
            return self.context_mappings[context_id]['matches'].get(target_statement, [])
        return []

    def get_period_for_context(self, context_id: str, statement_type: str) -> Optional[Tuple]:
        """Get period for a given context ID"""
        return self.fact_provider.filing_analyzer.statement_periods[statement_type].get(context_id)

    def get_matching_context(self, bs_context: str, statement_type: str) -> Optional[str]:
        """Get matching context for a statement type given a BS context"""
        mapping = self.context_mappings.get(bs_context, {})
        return mapping.get('matches', {}).get(statement_type)
        
    def get_valid_context_pairs(self) -> List[Dict[str, str]]:
        """Get all valid context pairs with their periods"""
        return [
            {
                'period': mapping['period'],
                'contexts': mapping['matches']
            }
            for mapping in self.context_mappings.values()
        ]

    def initialize_contexts(self):
        """Initialize context mappings after FilingAnalyzer has processed contexts"""
        self.context_mappings = self.fact_provider.filing_analyzer.context_mappings
        log_debug("\nInitialized FinancialDataManager context mappings from FilingAnalyzer (FinancialDataManager.initialize_contexts)")   

    def _create_context_mappings(self, statement_periods: dict) -> dict:
        """Create mappings between contexts with matching end dates"""
        mappings = {}
        log_debug("\n=== Creating Context Mappings ===")
        
        # Start with BS contexts as anchors
        for bs_context, bs_period in statement_periods['BS'].items():
            if not bs_period[1]:  # Skip if no end date
                continue
                
            log_debug(f"\nProcessing BS context: {bs_context}")
            log_debug(f"Period: {bs_period}")
            
            mappings[bs_context] = {
                'period': bs_period,
                'matches': {'IS': [], 'CI': [], 'CF': [], 'EQ': []}
            }
            
            # Find matching contexts in other statements
            for stmt_type in ['IS', 'CI', 'CF', 'EQ']:
                for other_context, other_period in statement_periods[stmt_type].items():
                    if other_period[1] == bs_period[1]:
                        mappings[bs_context]['matches'][stmt_type].append(other_context)
                        log_debug(f"Matched with {stmt_type} context: {other_context}")
            
            # Excessive Logging (remove me later)
            log_debug("\nDEBUG: Context Mappings Overview:")
            for bs_ctx, mapping in mappings.items():
                cf_matches = mapping['matches'].get('CF', [])
                log_debug(f"\nBS Context: {bs_ctx}")
                log_debug(f"Period: {mapping['period']}")
                log_debug(f"CF matches found: {len(cf_matches)}")
                if cf_matches:
                    for cf_ctx in cf_matches:
                        cf_period = self.statement_periods['CF'].get(cf_ctx)
                        if cf_period:
                            log_debug(f"  CF Context: {cf_ctx}")
                            log_debug(f"  Period: Start={cf_period[0].strftime('%Y-%m-%d') if cf_period[0] else 'None'}, End={cf_period[1].strftime('%Y-%m-%d')}")                        
        
        return mappings

    def find_matching_contexts(self) -> List[Tuple[str, str]]:
        """Find matching BS/IS contexts based on end dates"""
        log_debug("\n=== Finding Matching Contexts ===")
        
        matching_pairs = []
        
        # Add debug info about FilingAnalyzer instance
        log_debug(f"FilingAnalyzer instance id: {id(self.fact_provider.filing_analyzer)}")
        log_debug("Statement periods in FilingAnalyzer:")
        for stmt_type, periods in self.fact_provider.filing_analyzer.statement_periods.items():
            log_debug(f"- {stmt_type}: {len(periods)} periods")        
        
        # Get statement periods directly from FilingAnalyzer
        bs_statement = self.fact_provider.filing_analyzer.statement_periods.get('BS', {})
        is_statement = self.fact_provider.filing_analyzer.statement_periods.get('IS', {})
        
        if not bs_statement:
            log_debug("\nNo BS contexts found in FilingAnalyzer!")
            return []
        if not is_statement:
            log_debug("\nNo IS contexts found in FilingAnalyzer!")
            return []
                
        # Use statement_periods from FilingAnalyzer
        # bs_statement = self.statement_periods.get('BS', {})
        # is_statement = self.statement_periods.get('IS', {})
        
        if len(bs_statement.items()) == 0:
            log_debug("\nNo BS contexts found!")
            return []
        if len(is_statement.items()) == 0:
            log_debug("\nNo IS contexts found!")
            return []
            
        log_debug("\nAnalyzing BS contexts:")
        for bs_ctx, bs_period in bs_statement.items():
            if not bs_period or not bs_period[1]:
                continue
                
            log_debug(f"BS context {bs_ctx}: End={bs_period[1].strftime('%Y-%m-%d')}")
            
            for is_ctx, is_period in is_statement.items():
                if not is_period or not is_period[1]:
                    continue
                
                if (is_period[1] == bs_period[1] and 
                    is_period[0] is not None and 
                    bs_period[0] is None):
                    matching_pairs.append((bs_ctx, is_ctx))
                    log_debug(f"Found match - BS: {bs_ctx}, IS: {is_ctx}")
        
        log_debug(f"\nFound {len(matching_pairs)} matching context pairs")
        return matching_pairs

    def find_matching_cf_contexts(self, bs_context: str) -> List[str]:
        """Find CF contexts that match a BS context based on end dates"""
        log_debug(f"\n=== Finding Matching CF Contexts ===")
        log_debug(f"BS Context: {bs_context}")
        
        matching_cf_contexts = []
        
        try:
            # Get BS context period
            bs_period = self.fact_provider.filing_analyzer.statement_periods['BS'].get(bs_context)
            if not bs_period or not bs_period[1]:  # Need end date
                log_debug("No valid BS period found")
                return []
                
            bs_end_date = bs_period[1]
            log_debug(f"BS End Date: {bs_end_date.strftime('%Y-%m-%d')}")
            
            # Find CF contexts with matching end date
            for cf_context, cf_period in self.fact_provider.filing_analyzer.statement_periods['CF'].items():
                if not cf_period or not cf_period[1]:
                    log_debug(f"Skipping CF context {cf_context} - invalid period")
                    continue
                    
                cf_end_date = cf_period[1]
                if cf_end_date == bs_end_date:
                    # Safe period logging
                    period_str = (f"Period: {cf_period[0].strftime('%Y-%m-%d')} to {cf_period[1].strftime('%Y-%m-%d')}" 
                                if cf_period[0] else f"End date: {cf_period[1].strftime('%Y-%m-%d')}")
                    log_debug(f"Found matching CF context: {cf_context}")
                    log_debug(f"CF {period_str}")
                    matching_cf_contexts.append(cf_context)
                    
            if not matching_cf_contexts:
                log_debug("No matching CF contexts found")
                # Additional debug info
                log_debug("Available CF contexts and periods:")
                for ctx, period in self.fact_provider.filing_analyzer.statement_periods['CF'].items():
                    if period and period[1]:
                        end_date_str = period[1].strftime('%Y-%m-%d')
                        log_debug(f"  Context {ctx}: End date {end_date_str}")
                
            return matching_cf_contexts
            
        except Exception as e:
            log_debug(f"Error finding matching CF contexts: {str(e)}", include_trace=True)
            return []





class CalculationTracker:
    """
    Comprehensive tracker for financial calculations with audit trail.
    
    Features:
    - Calculation steps tracking
    - Input value recording
    - Component tracking
    - Time tracking
    - Audit trail
    """
    
    def __init__(self):
        self.calculations = []
        self.current_calculation = None
        self.ratio_summaries = defaultdict(list)
        self.logger = logger
             
          
    def add_input(self, name: str, value: float, metadata: ValueMetadata, 
                 is_component: bool = False):
        """Record an input value used in calculation"""
        if self.current_calculation:
            input_info = {
                'value': value,
                'metadata': metadata.to_dict() if metadata else None,
                'timestamp': datetime.now(),
                'is_component': is_component
            }
            
            if is_component:
                self.current_calculation['components'][name] = input_info
            else:
                self.current_calculation['inputs'][name] = input_info
                       
    def add_validation_result(self, is_valid: bool, messages: List[str]):
        """Record validation results"""
        if self.current_calculation:
            self.current_calculation['validation'] = {
                'is_valid': is_valid,
                'messages': messages,
                'timestamp': datetime.now()
            }
            
    def get_calculation_chain(self, identifier: str, period: tuple) -> List[dict]:
        """Get all calculations in a chain for a specific identifier"""
        return [
            calc for calc in self.calculations
            if calc['identifier'] == identifier and calc['period'] == period
        ]
        
    def get_last_calculation(self) -> Optional[dict]:
        """Get the most recent calculation details"""
        return self.calculations[-1] if self.calculations else None
        
    def get_calculation_stats(self) -> dict:
        """Get comprehensive calculation statistics"""
        stats = {
            'total_calculations': len(self.calculations),
            'completed': len([c for c in self.calculations if c['status'] == 'completed']),
            'errors': len([c for c in self.calculations if c['status'] == 'error']),
            'avg_duration': timedelta(0),
            'calculation_types': defaultdict(int),
            'input_usage': defaultdict(int)
        }
        
        if self.calculations:
            durations = [
                c['duration'] 
                for c in self.calculations 
                if 'duration' in c
            ]
            if durations:
                stats['avg_duration'] = sum(durations, timedelta(0)) / len(durations)
                
        # Count calculation types
        for calc in self.calculations:
            stats['calculation_types'][calc['type']] += 1
            # Track input usage
            for input_name in calc['inputs'].keys():
                stats['input_usage'][input_name] += 1
                
        return stats
        
    def get_audit_trail(self, calculation_type: Optional[str] = None) -> List[dict]:
        """
        Get detailed audit trail of calculations
        
        Args:
            calculation_type: Optional filter by calculation type
            
        Returns:
            List of calculation details sorted by timestamp
        """
        trail = []
        for calc in self.calculations:
            if calculation_type and calc['type'] != calculation_type:
                continue
                
            trail_entry = {
                'type': calc['type'],
                'identifier': calc['identifier'],
                'period': calc['period'],
                'start_time': calc['start_time'],
                'status': calc['status'],
                'result': calc.get('result'),
                'step_count': len(calc['steps']),
                'input_count': len(calc['inputs']),
                'component_count': len(calc['components']),
                'duration': calc.get('duration'),
                'validation': calc.get('validation')
            }
            trail.append(trail_entry)
            
        return sorted(trail, key=lambda x: x['start_time'])
        
    def clear(self):
        """Clear all tracked calculations"""
        self.calculations.clear()
        self.current_calculation = None
        
    def export_calculation_history(self) -> dict:
        """Export full calculation history in a structured format"""
        return {
            'calculations': self.calculations,
            'stats': self.get_calculation_stats(),
            'audit_trail': self.get_audit_trail(),
            'export_time': datetime.now().isoformat()
        }

    def track_value_access(self, concept: str, context_id: str, value: float, metadata: dict):
        """Track value access with context awareness"""
        log_debug(f"\nTracking value access: {concept} (context: {context_id})")
        
        if not self.current_calculation:
            log_debug("No active calculation - skipping tracking")
            return
            
        input_info = {
            'value': value,
            'context_id': context_id,
            'metadata': metadata,
            'timestamp': datetime.now()
        }
        
        self.current_calculation['inputs'][concept] = input_info
        log_debug(f"Added input tracking for {concept}")

    def track_ratio_calculation(self, ratio_name: str, ratio_value: float, 
                              context_ids: Dict[str, str], 
                              inputs_used: dict):
        """Track ratio calculation with basic metrics"""
        log_debug(f"\nTracking calculation for {ratio_name}")
        log_debug(f"Result: {ratio_value:.4f}")
        log_debug(f"Context IDs: {context_ids}")
        
        calculation_info = {
            'ratio_name': ratio_name,
            'result': ratio_value,
            'context_ids': context_ids,
            'inputs': inputs_used,
            'timestamp': datetime.now()
        }
        self.ratio_summaries[ratio_name].append(calculation_info)

    def summarize_calculations(self):
        """Generate comprehensive calculation summary"""
        log_debug("\n=== Calculation Summary ===")
        
        # Summarize ratio calculations
        if self.ratio_summaries:
            log_debug("\nRatio Calculations:")
            for ratio_name, calcs in self.ratio_summaries.items():
                log_debug(f"\n{ratio_name}:")
                log_debug(f"Total calculations: {len(calcs)}")
                
                # Context usage analysis
                bs_contexts = {c['context_ids'].get('BS') for c in calcs}  # Changed from 'contexts'
                is_contexts = {c['context_ids'].get('IS') for c in calcs}  # Changed from 'contexts'
                bs_contexts.discard(None)
                is_contexts.discard(None)
                
                if bs_contexts:
                    log_debug(f"BS contexts used: {sorted(bs_contexts)}")
                if is_contexts:
                    log_debug(f"IS contexts used: {sorted(is_contexts)}")
                    
                # Value range analysis
                values = [c['result'] for c in calcs]
                if values:
                    log_debug(f"Value range: {min(values):.4f} to {max(values):.4f}")
                    
                # Input analysis
                all_inputs = set()
                for calc in calcs:
                    all_inputs.update(calc['inputs'].keys())
                
                if all_inputs:
                    log_debug("Inputs required:")
                    for input_name in sorted(all_inputs):
                        usage_count = sum(1 for c in calcs if input_name in c['inputs'])
                        log_debug(f"  - {input_name}: used in {usage_count}/{len(calcs)} calculations")
            ###################################                        
            # Overall statistics
            log_debug("\nOverall Statistics:")
            total_ratios = sum(len(calcs) for calcs in self.ratio_summaries.values())
            unique_ratios = len(self.ratio_summaries)
            log_debug(f"Total ratio calculations: {total_ratios}")
            log_debug(f"Unique ratio types: {unique_ratios}")
            
            # Success rate analysis
            successful = sum(1 for calcs in self.ratio_summaries.values() 
                        for calc in calcs if calc['result'] is not None)
            if total_ratios > 0:
                success_rate = (successful / total_ratios) * 100
                log_debug(f"Success rate: {success_rate:.1f}%")
        
        # Log other calculations
        if self.calculations:
            log_debug(f"\nOther calculations: {len(self.calculations)}")
            calculation_types = Counter(c['type'] for c in self.calculations)
            for calc_type, count in calculation_types.items():
                log_debug(f"  {calc_type}: {count}")
                
                # Show error summary if any
                errors = [c for c in self.calculations if c['type'] == calc_type and c.get('status') == 'error']
                if errors:
                    log_debug(f"  Errors in {calc_type}:")
                    for error in errors:
                        log_debug(f"    - {error.get('error', 'Unknown error')}")

    def add_step(self, description: str, value: Optional[float] = None, 
                metadata: Optional[Dict] = None):
        """Record a calculation step"""
        if self.current_calculation:
            step = {
                'description': description,
                'value': value,
                'metadata': metadata,
                'timestamp': datetime.now()
            }
            self.current_calculation['steps'].append(step)

    def start_calculation(self, calculation_type: str, identifier: str, context_id: str):
        """Start tracking a new calculation"""
        if self.current_calculation:
            self.complete_calculation(None)
            
        self.current_calculation = {
            'type': calculation_type,
            'identifier': identifier,
            'context_id': context_id,
            'steps': [],
            'inputs': {},
            'components': {},
            'start_time': datetime.now(),
            'status': 'in_progress'
        }
        
        self.calculations.append(self.current_calculation)
        log_debug(f"Started tracking calculation: {identifier} (context: {context_id})")

    def complete_calculation(self, result: Optional[float], 
                           status: str = 'completed', 
                           error: Optional[str] = None):
        """Complete the current calculation"""
        if self.current_calculation:
            self.current_calculation.update({
                'result': result,
                'end_time': datetime.now(),
                'duration': datetime.now() - self.current_calculation['start_time'],
                'status': status,
                'error': error
            })
            self.current_calculation = None



        
class FactProvider:
    """Provides standardized access to financial facts across taxonomies"""
    def __init__(self, model_xbrl, existing_analyzer=None):
        self.model_xbrl = model_xbrl
        self.logger = logger
        self.custom_taxonomies_used = set()
        self.concept_maps = FinancialConceptMaps.get_concept_maps()
        self.filing_analyzer = existing_analyzer or FilingAnalyzer(model_xbrl)
        #log_debug(f"\nFactProvider using FilingAnalyzer instance: {id(self.filing_analyzer)}")
        log_debug("\nInitializing FactProvider with FinancialDataManager")
        self.financial_data_manager = FinancialDataManager(self)
        log_debug("FinancialDataManager initialized")
    
    def get_fact_value_by_context(self, concept_key: str, context_id: str, statement_type: str) -> tuple:
        """Internal method to get fact value directly by context"""
        log_debug(f"\nLooking up: {concept_key} (context: {context_id})")
        
        try:
            concept_info = self.concept_maps.get(concept_key)
            if not concept_info:
                return None, None, False

            # Direct taxonomy lookups only - no compound calculations
            if 'taxonomies' in concept_info:
                for taxonomy, concepts in concept_info['taxonomies'].items():
                    for concept_name in concepts:
                        qname = f"{taxonomy}:{concept_name}"
                        value = self._find_fact_value(qname, context_id)
                        if value is not None:
                            is_custom = taxonomy not in ['ifrs-full', 'us-gaap']
                            log_debug(f"Found value {value} ({qname})")
                            return value, qname, is_custom

            return None, None, False

        except Exception as e:
            self.logger.log(f"Error in get_fact_value_by_context: {str(e)}", include_trace=True)
            return None, None, False

    def _find_fact_value(self, qname: str, context_id: str) -> Optional[float]:
        """Internal helper to find a single fact value"""
        try:
            # log_debug(f"\n=== Looking up fact value: {qname} for contextRef {context_id} ===")
            # Get facts by qname
            if isinstance(qname, str):
                prefix, localname = qname.split(':')
                facts = [f for f in self.model_xbrl.facts 
                        if f.qname.prefix == prefix and f.qname.localName == localname]
                # log_debug(f"Found {len(facts)} facts for {qname}")
            else:
                facts = self.model_xbrl.factsByQname.get(qname, [])
                # log_debug(f"Found {len(facts)} facts using direct qname lookup")

            # Find matching fact by context
            for fact in facts:
                if fact.context and fact.context.id == context_id:
                    try:
                        value = float(fact.value)
                        # log_debug(f"Found value {value} for context {context_id}")
                        return value
                    except (ValueError, TypeError):
                        continue

            # log_debug(f"No matching fact found for context {context_id}")
            return None

        except Exception as e:
            self.logger.log(f"Error finding fact value: {str(e)}", include_trace=True)
            return None

    def _track_fact_lookup(self, concept_key: str, context_id: str, statement_type: str):
        """Track fact lookup attempts for debugging"""
        log_debug(f"\nTracking fact lookup:")
        log_debug(f"  Concept: {concept_key}")
        log_debug(f"  Context: {context_id}")
        log_debug(f"  Statement Type: {statement_type}")
        
        # Check if concept mapping exists
        concept_info = self.concept_maps.get(concept_key)
        if concept_info:
            log_debug(f"  Found concept mapping with {len(concept_info['taxonomies'])} taxonomies")
        else:
            log_debug("  No concept mapping found")
   
    
from typing import Dict, List, Tuple, DefaultDict
from collections import defaultdict    
class RatioInputTracker:
    """
    Tracks and formats ratio calculation inputs with improved tracking
    and component handling.
    """
    
    def __init__(self, data_manager):
        self.data_manager = data_manager
        self.inputs = defaultdict(list)
        self.group_order = [
            ['total_assets', 'total_liabilities', 'total_equity'],
            ['revenue', 'net_income'],
            ['total_debt'],
            ['total_capital'],
            ['quick_assets'],              
            ['gross_profit'],
            ['cash_and_equivalents'],
            ['net_debt'],            
            ['working_capital'],
            ['ebit', 'ebitda', 'ebt'],
            ['operating_expenses', 'depreciation_amortization', 'interest_expense', 'income_tax_expense'],
            ['operating_cashflow'],
            ['inventory', 'cost_of_sales'],  # Inventory Turnover
            ['dividends', 'net_income'],  # Dividend Payout            
            ['long_term_borrowings']  # Added long-term borrowings
        ]

    def track_ratio_calculation(self, ratio_name: str, context_ids: Dict[str, str], 
                            inputs_used: dict, component_details: dict = None):
        """Track ratio calculation inputs with context awareness"""
        log_debug(f"\nTracking inputs for {ratio_name}")
        log_debug(f"Using contexts: {context_ids}")
        log_debug(f"Inputs used: {inputs_used}")
        
        # Track main inputs
        for concept, (value, qname, is_custom) in inputs_used.items():
            context_id = context_ids.get('BS' if concept in ['total_assets', 'total_liabilities','quick_assets', 'total_debt','cash_and_equivalents', 'net_debt'] else 'IS')
            if context_id:
                self.add_input(
                    name=concept,
                    qname=qname, 
                    value=value,
                    context_id=context_id,
                    statement=self._determine_statement_type(concept),
                    indent_level=0
                )
                log_debug(f"Added input: {concept} = {value} for context {context_id}")
                    
        # Track components 
        if component_details:
            log_debug(f"Processing component details: {component_details}")
            self._track_components(component_details, context_ids)

    def _track_components(self, component_details: dict, period: tuple, indent_level: int = 1):
        """Recursively track component values"""
        for component, details in component_details.get('components', {}).items():
            sign = '+' if details.get('value', 0) >= 0 else '-'
            self.add_input(
                name=component,
                qname=details.get('metadata', {}).get('qname', 'Unknown'),
                value=abs(details.get('value', 0)),
                period=period,
                statement=details.get('metadata', {}).get('source_statement', 'Unknown'),
                indent_level=indent_level
                # sign=sign
            )
            
            # Recurse for nested components
            if 'components' in details:
                self._track_components(details, period, indent_level + 1)

    def add_input(self, name: str, qname: str, value: float, context_id: str, 
                statement: str, indent_level: int = 0, calculation_rule: str = None):
        """Add a single input value with contextRef tracking"""
        log_debug(f"\nAdding input - Name: {name}, Value: {value}, Context: {context_id}, Indent: {indent_level}")
        
        if value is None or context_id is None:
            log_debug("Skipping due to None value or context")
            return
            
        # Get calculation rule for parent if this is a component
        if indent_level > 0 and calculation_rule is None and context_id in self.inputs:
            # Find parent component
            parent_entries = [x for x in self.inputs[context_id] if x['indent'] == indent_level - 1]
            if parent_entries:
                parent_name = parent_entries[-1]['name']
                calculation_rule = self.data_manager.component_registry.get_calculation_rule(parent_name)
                log_debug(f"Found parent {parent_name} with calculation rule: {calculation_rule}")

        # Determine sign for display
        if indent_level > 0:
            # For subtraction components after first component, show negative sign
            is_first_component = not any(x for x in self.inputs[context_id] 
                                       if x['indent'] == indent_level and x['is_component'])
            # display_sign = "+" if (is_first_component or calculation_rule != "subtract") else "-"                                 # with sign
            display_sign = " "                                                                                                   # without sign
            # display_name = '  ' * indent_level + f"{display_sign} {name}"                                                       # with sign
            display_name = '  ' * indent_level + name                                                                           # without sign
            log_debug(f"Component sign for {name}: {display_sign} (first: {is_first_component}, rule: {calculation_rule})")
        else:
            display_name = name
            
        # Check if this input already exists to avoid duplicates
        existing = next((x for x in self.inputs[context_id] 
                        if x['name'] == name and x['indent'] == indent_level), None)
        
        if existing:
            log_debug(f"Updating existing input: {name}")
            existing.update({
                'value': value,
                'qname': qname,
                'statement': statement
            })
        else:
            log_debug(f"Adding new input: {display_name}")
            input_data = {
                'name': name,
                'display_name': display_name,
                'qname': qname,
                'value': value,
                'statement': statement,
                'indent': indent_level,
                'is_component': indent_level > 0,
                'context_id': context_id
            }
            self.inputs[context_id].append(input_data)
            
            # Track recursive components if it's a compound value
            if indent_level > 0 and value is not None:
                if self.data_manager.component_registry._registry.get(name):
                    compound_value, compound_metadata = self.data_manager.get_value_by_context(
                        name, context_id, statement)
                    if compound_value is not None and hasattr(compound_metadata, 'components'):
                        comp_calc_rule = self.data_manager.component_registry.get_calculation_rule(name)
                        for comp_name, comp_metadata in compound_metadata.components.items():
                            comp_value = getattr(comp_metadata, 'value', None)  
                            comp_qname = getattr(comp_metadata, 'qname', 'Unknown')
                            if comp_value is not None:
                                self.add_input(
                                    name=comp_name,
                                    qname=comp_qname,
                                    value=comp_value,
                                    context_id=context_id,
                                    statement=statement,
                                    indent_level=indent_level + 1,
                                    calculation_rule=comp_calc_rule
                                )
            
        log_debug(f"Successfully handled input for {context_id}")

    def _determine_statement_type(self, concept: str) -> str:
        """Determine statement type for a concept"""
        bs_concepts = {'total_assets', 'total_liabilities', 'total_equity', 
                    'total_debt', 'total_capital', 'working_capital', 'quick_assets',
                    'cash_and_equivalents', 'inventory', 'current_liabilities','net_debt', 'long_term_borrowings'}  # Added inventory
                    
        # All our new concepts are either IS or CF concepts
        is_concepts = {'revenue', 'net_income', 'ebit', 'ebitda', 'ebt', 
                    'operating_expenses', 'depreciation_amortization', 
                    'interest_expense', 'income_tax_expense',
                    'cost_of_sales', 'dividends'}  # Added cost_of_sales, dividends
                    
        cf_concepts = {'operating_cashflow', 'dividends'}  # Added CF concept
        
        if concept in bs_concepts:
            return 'BS'
        elif concept in cf_concepts:
            return 'CF'
        else:
            return 'IS'  # Default to IS for other concepts

    def get_organized_inputs(self, context_pairs: List[Tuple[str, str]]) -> List[List]:
        """Get organized inputs by groups with proper statement type detection"""
        log_debug("\n=== Organizing Ratio Inputs ===")
        rows = []
        pair_count = len(context_pairs)
        log_debug(f"Processing {pair_count} context pairs")
        
        for group in self.group_order:
            log_debug(f"\nProcessing group: {group}")
            group_data = []
            
            for metric in group:
                log_debug(f"\nProcessing metric: {metric}")
                metric_rows = {'main': []}
                components = defaultdict(list)
                
                # Check if this is a compound value
                is_compound = metric in self.data_manager.component_registry._registry
                
                # Get statement type from concept maps
                concept_info = self.data_manager.fact_provider.concept_maps.get(metric)
                if not concept_info:
                    log_debug(f"No concept mapping found for {metric} - skipping")
                    continue
                    
                stmt_type = concept_info['statements'][0]  # Use first statement type listed
                log_debug(f"Found statement type {stmt_type} for {metric}")
                
                # Process each context pair
                for bs_ctx, is_ctx in context_pairs:
                    context_id = bs_ctx if stmt_type == 'BS' else is_ctx
                    
                    # Get main value
                    value, metadata = self.data_manager.get_value_by_context(metric, context_id, stmt_type)
                    if value is not None:
                        log_debug(f"Found value for {metric}: {value}")
                        main_entry = {
                            'name': metric,
                            'display_name': metric,
                            'qname': metadata.qname if metadata else 'Unknown',
                            'value': value,
                            'statement': metadata.source_statement if metadata else stmt_type,
                            'indent': 0,
                            'is_component': False,
                            'context_id': context_id
                        }
                        metric_rows['main'].append((context_id, main_entry))
                        log_debug("Added main entry")
                        
                        # If this is a compound value, get its components
                        if is_compound:
                            log_debug(f"Processing components for compound value {metric}")
                            compound_components = self.data_manager.component_registry.get_components(metric)
                            calculation_rule = self.data_manager.component_registry.get_calculation_rule(metric)
                            
                            for i, component in enumerate(compound_components):
                                # Get statement type for component
                                comp_concept_info = self.data_manager.fact_provider.concept_maps.get(component)
                                if comp_concept_info:
                                    comp_stmt_type = comp_concept_info['statements'][0]
                                    comp_context_id = bs_ctx if comp_stmt_type == 'BS' else is_ctx
                                    
                                    comp_value, comp_metadata = self.data_manager.get_value_by_context(
                                        component, comp_context_id, comp_stmt_type)
                                        
                                    if comp_value is not None:
                                        # Determine sign based on calculation rule and position
                                        sign = ''  # Default no sign for first component
                                        if i > 0:  # After first component
                                            if calculation_rule == "subtract":
                                                sign = "-"
                                            elif calculation_rule == "sum":
                                                sign = "+"
                                            elif calculation_rule == "divide":
                                                sign = "/" if i == 1 else "+"
                                                
                                        log_debug(f"Found component {component}: {comp_value} (sign: {sign})")
                                        display_name = f"  {sign} {component}" if sign else f"  {component}"
                                        
                                        components[display_name].append((comp_context_id, {
                                            'name': component,
                                            'display_name': display_name,
                                            'qname': comp_metadata.qname if comp_metadata else 'Unknown',
                                            'value': comp_value,
                                            'statement': comp_metadata.source_statement if comp_metadata else comp_stmt_type,
                                            'indent': 1,
                                            'is_component': True,
                                            'context_id': comp_context_id
                                        }))
                
                if metric_rows['main']:
                    # Add main row
                    main_data = metric_rows['main'][0][1]
                    row = [
                        main_data['display_name'],
                        main_data['statement'],
                        main_data['qname']
                    ] + [''] * pair_count
                    
                    # Fill values
                    for pair_idx, (bs_ctx, is_ctx) in enumerate(context_pairs):
                        # Use statement type from concept map
                        ctx = bs_ctx if stmt_type == 'BS' else is_ctx
                        match = next((m[1] for m in metric_rows['main'] if m[0] == ctx), None)
                        if match:
                            row[3 + pair_idx] = match['value']
                            
                    group_data.append(row)
                    
                    # Add component rows
                    for comp_name, comp_contexts in components.items():
                        if comp_contexts:
                            comp_data = comp_contexts[0][1]
                            comp_row = [
                                comp_data['display_name'],
                                comp_data['statement'],
                                comp_data['qname']
                            ] + [''] * pair_count
                            
                            for pair_idx, (bs_ctx, is_ctx) in enumerate(context_pairs):
                                comp_concept = comp_data['name']
                                comp_concept_info = self.data_manager.fact_provider.concept_maps.get(comp_concept)
                                if comp_concept_info:
                                    comp_stmt_type = comp_concept_info['statements'][0]
                                    ctx = bs_ctx if comp_stmt_type == 'BS' else is_ctx
                                    match = next((c[1] for c in comp_contexts if c[0] == ctx), None)
                                    if match:
                                        comp_row[3 + pair_idx] = match['value']
                                        
                            group_data.append(comp_row)
            
            if group_data:
                rows.extend(group_data)
                rows.append([''] * (3 + pair_count))
                log_debug(f"Added {len(group_data)} rows for group")
        
        log_debug(f"\nFinal output: {len(rows)} rows")
        return rows




class RatioResultFormat:
    """Standardizes ratio result format and provides validation"""
    
    # Required fields for valid ratio results
    REQUIRED_FIELDS = {
        'ratio_name': str,
        'value': float,
        'period': tuple,
        'context_ids': dict,  # Must contain at least one valid contextRef
        'inputs_used': dict   # Must contain at least one input
    }
    
    @staticmethod
    def validate(result: 'RatioResult') -> Tuple[bool, List[str]]:
        """Validate ratio result format"""
        errors = []
        
        # Check required fields
        for field, field_type in RatioResultFormat.REQUIRED_FIELDS.items():
            if not hasattr(result, field):
                errors.append(f"Missing required field: {field}")
                continue
                
            value = getattr(result, field)
            if not isinstance(value, field_type):
                errors.append(f"Invalid type for {field}: expected {field_type}, got {type(value)}")
                
        # Validate period format
        if hasattr(result, 'period'):
            if not (isinstance(result.period, tuple) and len(result.period) == 2):
                errors.append("Invalid period format: must be tuple(start_date, end_date)")
            elif result.period[1] is None:
                errors.append("Invalid period: end_date cannot be None")
                
        # Validate context_ids
        if hasattr(result, 'context_ids'):
            if not result.context_ids:
                errors.append("context_ids cannot be empty")
            elif not all(isinstance(k, str) and isinstance(v, str) 
                        for k, v in result.context_ids.items()):
                errors.append("Invalid context_ids format")
                
        # Validate inputs_used
        if hasattr(result, 'inputs_used'):
            if not result.inputs_used:
                errors.append("inputs_used cannot be empty")
            for input_name, input_data in result.inputs_used.items():
                if not isinstance(input_data, tuple) or len(input_data) != 3:
                    errors.append(f"Invalid input format for {input_name}")
                    
        return len(errors) == 0, errors
    
    
class RatioResult:
    """Enhanced container for ratio calculation results with context-based tracking"""
    
    def __init__(self, ratio_name: str, value: float, context_ids: Dict[str, str],
                 inputs_used: dict, calculation_variant: str = None,
                 component_details: dict = None):
        """
        Args:
            ratio_name: Name of the ratio
            value: Calculated ratio value
            context_ids: Dict mapping statement types to context IDs {'BS': 'ctx1', 'IS': 'ctx2'}
            inputs_used: Dict of inputs used with their metadata
            calculation_variant: Optional variant of calculation method used
            component_details: Optional details about components used
        """
        self.ratio_name = ratio_name
        self.value = float(value)  # Ensure float type
        self.context_ids = context_ids
        self.inputs_used = inputs_used
        self.calculation_variant = calculation_variant
        self.component_details = component_details
        self.calculation_time = datetime.now()
        
        # Validate on creation
        is_valid, errors = self._validate()
        if not is_valid:
            log_debug("WARNING: Invalid ratio result format:")
            for error in errors:
                log_debug(f"- {error}")
 
    def _validate(self) -> Tuple[bool, List[str]]:
        """
        Validate ratio result format with enhanced context validation
        using concept maps as source of truth
        """
        errors = []
        
        # Required fields
        if not self.ratio_name or not isinstance(self.ratio_name, str):
            errors.append("Invalid ratio_name")
            
        if not isinstance(self.value, (int, float)):
            errors.append("Invalid value type")
            
        # Context validation
        if not self.context_ids:
            errors.append("context_ids cannot be empty")
        elif not isinstance(self.context_ids, dict):
            errors.append("context_ids must be a dictionary")
        else:
            # Get concept maps from fact provider to check statement requirements
            concept_maps = FinancialConceptMaps.get_concept_maps()
            
            # Collect required statement types based on inputs used
            required_statements = set()
            for input_name in self.inputs_used.keys():
                if input_name in concept_maps:
                    # Get statements list from concept map
                    statements = concept_maps[input_name].get('statements', [])
                    required_statements.update(statements)
                    log_debug(f"Input {input_name} requires statements: {statements}")
                else:
                    log_debug(f"Warning: Input {input_name} not found in concept maps")
            
            # Validate context requirements based on used inputs
            missing_contexts = required_statements - set(self.context_ids.keys())
            if missing_contexts:
                errors.append(f"Missing required contexts for {self.ratio_name} based on inputs: {missing_contexts}")
                
            # Validate that all provided contexts correspond to valid statement types
            valid_types = {'BS', 'IS', 'CF', 'CI', 'EQ'}  # Keep this as a basic validation
            invalid_types = set(self.context_ids.keys()) - valid_types
            if invalid_types:
                errors.append(f"Invalid statement types in context_ids: {invalid_types}")
        
        # Input validation
        if not self.inputs_used:
            errors.append("inputs_used cannot be empty")
        else:
            for input_name, input_data in self.inputs_used.items():
                if not isinstance(input_data, tuple) or len(input_data) != 3:
                    errors.append(f"Invalid input format for {input_name}")
                
                # Check if input has matching context
                if input_name in concept_maps:
                    # Get primary statement type from concept map
                    primary_stmt = concept_maps[input_name]['statements'][0] if concept_maps[input_name]['statements'] else None
                    if primary_stmt and primary_stmt not in self.context_ids:
                        errors.append(f"Missing {primary_stmt} context for input {input_name}")
                
        return len(errors) == 0, errors

    def _get_required_contexts(self) -> Set[str]:
        """Determine required contexts based on ratio type"""
        # Define context requirements for each ratio type
        requirements = {
            'Return on Assets': {'BS', 'IS'},
            'Current Ratio': {'BS'},
            'Debt to Equity Ratio': {'BS'},
            'Net Margin': {'IS'},
            'Gross Margin': {'IS'},
            'Cost of Sales Ratio': {'IS'}
        }
        return requirements.get(self.ratio_name, {'BS'})  # Default to BS if unknown

    def get_formatted_period(self, data_manager) -> str:
        """Get formatted period string using primary context"""
        try:
            # Use BS context if available, otherwise first available context
            context_id = self.context_ids.get('BS') or next(iter(self.context_ids.values()))
            stmt_type = 'BS' if 'BS' in self.context_ids else next(iter(self.context_ids.keys()))
            
            period = data_manager.get_period_for_context(context_id, stmt_type)
            if period:
                if period[0] is None:
                    return f"As of {period[1].strftime('%Y-%m-%d')}"
                return f"{period[0].strftime('%Y-%m-%d')} - {period[1].strftime('%Y-%m-%d')}"
            return "Unknown Period"
        except Exception as e:
            log_debug(f"Error formatting period: {str(e)}")
            return "Unknown Period"

    def get_inputs_summary(self) -> str:
        """Get formatted summary of inputs with context references"""
        try:
            parts = []
            for name, (value, _, _) in self.inputs_used.items():
                stmt_type = 'BS' if name in ['total_assets', 'total_liabilities', 'current_assets', 'current_liabilities'] else 'IS'
                context_id = self.context_ids.get(stmt_type, 'unknown')
                parts.append(f"{name}: {value:.2f} ({context_id})")
            return "; ".join(parts)
        except Exception as e:
            log_debug(f"Error creating inputs summary: {str(e)}")
            return "Error summarizing inputs"

    def to_dict(self) -> dict:
        """Convert to dictionary format for storage/export"""
        return {
            'ratio': self.ratio_name,
            'value': self.value,
            'context_ids': self.context_ids,
            'inputs': {
                name: {
                    'value': value,
                    'qname': qname,
                    'is_custom': is_custom,
                    'context_id': self.context_ids.get(
                        'BS' if name in ['total_assets', 'total_liabilities'] else 'IS'
                    )
                }
                for name, (value, qname, is_custom) in self.inputs_used.items()
            },
            'calculation': {
                'variant': self.calculation_variant,
                'components': self.component_details,
                'timestamp': self.calculation_time.isoformat()
            }
        }

    def add_component_metadata(self, component_name: str, metadata: ValueMetadata) -> None:
        """Track metadata for calculated components"""
        if self.component_details is None:
            self.component_details = {}
        
        self.component_details[component_name] = {
            'value': metadata.value if hasattr(metadata, 'value') else None,
            'qname': metadata.qname,
            'source_statement': metadata.source_statement,
            'is_custom': metadata.is_custom,
            'context_id': metadata.context_id,
            'last_updated': metadata.last_updated.isoformat()
        }
        log_debug(f"Added component metadata for {component_name}")


class RatioResultSet:
    """Container for ratio results associated with specific contexts"""
    
    def __init__(self, bs_context: str, data_manager):
        self.bs_context = bs_context
        self.data_manager = data_manager
        self.results = {}  # ratio_name -> RatioResult
        self.calculation_time = datetime.now()
        self.format_errors = []
        
        log_debug(f"\nInitialized RatioResultSet for BS context: {bs_context}")
        
    def add_result(self, result: RatioResult) -> bool:
        """Add ratio result with enhanced validation"""
        log_debug(f"\nAdding result: {result.ratio_name}")
        log_debug(f"Inputs used: {result.inputs_used}")
        
        try:
            # First check basic result validity
            is_valid, errors = result._validate()
            if not is_valid:
                log_debug("Validation failed:")
                for error in errors:
                    log_debug(f"- {error}")
                self.format_errors.append({
                    'ratio': result.ratio_name,
                    'errors': errors,
                    'timestamp': datetime.now()
                })
                return False

            # Verify BS context consistency
            if 'BS' in result.context_ids:
                if result.context_ids['BS'] != self.bs_context:
                    log_debug(f"Context mismatch: Expected BS context {self.bs_context}, got {result.context_ids['BS']}")
                    return False
            if 'CF' in result.context_ids:
                if result.context_ids['CF'] not in self.data_manager.fact_provider.filing_analyzer.statement_periods['CF']:
                    log_debug(f"Invalid CF context: {result.context_ids['CF']}")
                    return False            
                    
            # Store result
            self.results[result.ratio_name] = result
            log_debug(f"Successfully added {result.ratio_name} result")
            return True
                
        except Exception as e:
            log_debug(f"Error adding result: {str(e)}")
            return False
        
    def get_context_periods(self) -> Dict[str, tuple]:
        """Get periods for all contexts in this set"""
        periods = {}
        for result in self.results.values():
            for stmt_type, context_id in result.context_ids.items():
                if context_id not in periods:
                    period = self.data_manager.get_period_for_context(context_id, stmt_type)
                    if period:
                        periods[context_id] = period
        return periods
               
    def get_result(self, ratio_name: str) -> Optional[RatioResult]:
        """Get a specific ratio result"""
        return self.results.get(ratio_name)
        
    def to_dict(self) -> dict:
        """Convert entire result set to dictionary"""
        return {
            'bs_context': self.bs_context,
            'period': self.data_manager.get_period_for_context(self.bs_context, 'BS'),
            'results': {
                name: result.to_dict()
                for name, result in self.results.items()
            },
            'calculation_time': self.calculation_time.isoformat()
        }

    def get_validation_status(self) -> dict:
        """Get validation status for all results"""
        return {
            'bs_context': self.bs_context,
            'valid_results': len(self.results),
            'invalid_results': len(self.format_errors),
            'errors': self.format_errors
        }

    def add_to_analysis(self, analysis: 'RatioAnalysis'):
        """Add results to ratio analysis"""
        analysis.result_sets[self.bs_context] = self
        log_debug(f"Added {len(self.results)} results to analysis for context {self.bs_context}")    
  
    def aggregate_by_period(self) -> Dict[str, List[RatioResult]]:
        """Group results by reporting period"""
        log_debug("\nAggregating results by period")
        
        period_groups = defaultdict(list)
        for ratio_name, result in self.results.items():
            period_str = result.get_formatted_period(self.data_manager)
            period_groups[period_str].append(result)
            
        log_debug(f"Created {len(period_groups)} period groups")
        return dict(period_groups)

    def get_component_summary(self) -> Dict[str, Dict]:
        """Get summary of all components used in calculations"""
        log_debug("\nGenerating component summary")
        
        components = {}
        for result in self.results.values():
            if result.component_details:
                for comp_name, comp_data in result.component_details.items():
                    if comp_name not in components:
                        components[comp_name] = {
                            'usage_count': 0,
                            'ratios': set(),
                            'is_custom': comp_data.get('is_custom', False)
                        }
                    components[comp_name]['usage_count'] += 1
                    components[comp_name]['ratios'].add(result.ratio_name)
        
        log_debug(f"Found {len(components)} unique components")
        return components  
             
    
class RatioAnalysis:
    """Comprehensive ratio analysis with trends and metadata"""
    
    def __init__(self, model_xbrl):
        self.model_xbrl = model_xbrl
        self.result_sets = {}  # period -> RatioResultSet
        self.data_manager = model_xbrl._fact_provider.financial_data_manager
        self.metadata = {
            'calculation_time': datetime.now(),
            'ratio_count': 0,
            'period_count': 0
        }
        
    def add_result_set(self, result_set: RatioResultSet):
        """Add a result set with enhanced validation"""
        if not result_set or not result_set.bs_context:
            log_debug("Invalid result set - skipping")
            return
            
        log_debug(f"\nAdding result set for context: {result_set.bs_context}")
        
        # Validate results before adding
        validation = result_set.get_validation_status()
        if validation['invalid_results'] > 0:
            log_debug(f"Warning: {validation['invalid_results']} invalid results found")
            for error in validation['errors']:
                log_debug(f"- {error['ratio']}: {', '.join(error['errors'])}")
        
        self.result_sets[result_set.bs_context] = result_set
        self._update_metadata()
        
        log_debug(f"Added {validation['valid_results']} valid ratios")
    
    def _update_metadata(self):
        """Update analysis metadata"""
        self.metadata.update({
            'ratio_count': len(self.get_ratio_names()),
            'period_count': len(self.result_sets),
            'last_update': datetime.now()
        })
        
    def get_ratio_names(self) -> Set[str]:
        """Get all unique ratio names"""
        return {
            ratio_name
            for result_set in self.result_sets.values()
            for ratio_name in result_set.results.keys()
        }
        
    def get_ratio_trend(self, ratio_name: str) -> List[Tuple[tuple, float]]:
        """Get trend data for a specific ratio"""
        return [
            (period, result_set.results[ratio_name].value)
            for period, result_set in sorted(self.result_sets.items())
            if ratio_name in result_set.results
        ]
        
    def prepare_excel_data(self) -> dict:
        """Prepare ratio data for Excel export with context awareness"""
        log_debug("\n=== Preparing Ratio Data ===")
        
        headers = [
            'Ratio Name', 'Value', 'Period', 'Statement Types Used',
            'Calculation Method', 'Inputs Used', 'Data Quality'
        ]
        data = []
        
        try:
            # Get context groups first - using existing get_context_groups
            context_groups = self.get_context_groups()
            log_debug(f"Processing {len(context_groups)} context groups")
            
            for period_str, results in context_groups.items():
                # Add group separator
                if data:
                    data.append([''] * len(headers))
                    
                log_debug(f"\nProcessing group: {period_str}")
                
                for result in results:
                    # Get statement types used
                    stmt_types = sorted(set(result.context_ids.keys()))
                    
                    # Create row with enhanced information
                    row = [
                        result.ratio_name,
                        result.value,
                        period_str,
                        ' & '.join(stmt_types),
                        result.calculation_variant or 'standard',
                        result.get_inputs_summary(),
                        self._assess_data_quality(result)
                    ]
                    data.append(row)
                    log_debug(f"Added {result.ratio_name}: {result.value:.4f}")
            
            log_debug(f"\nPrepared {len(data)} ratio records for export")
            return {'headers': headers, 'data': data}
            
        except Exception as e:
            log_debug(f"Error preparing ratio data: {str(e)}", include_trace=True)
            return {'headers': headers, 'data': []}
       
    def _format_context_period(self, context_id: str) -> str:
        """Format period based on context"""
        period = self.data_manager.get_period_for_context(context_id, 'BS')
        if period[0] is None:
            return f"As of {period[1].strftime('%Y-%m-%d')}"
        return f"{period[0].strftime('%Y-%m-%d')} - {period[1].strftime('%Y-%m-%d')}"

    def _format_inputs_summary(self, result: RatioResult) -> str:
        """Format inputs summary with context references"""
        summaries = []
        for name, (value, qname, is_custom) in result.inputs_used.items():
            context_id = result.context_ids.get(
                'BS' if name in ['total_assets', 'total_liabilities'] else 'IS'
            )
            if context_id:
                summaries.append(f"{name}: {value:.2f} ({context_id})")
        return "; ".join(summaries)

    def add_ratio_result(self, result: RatioResult):
        """Add ratio result with proper period tracking"""
        period = result.period
        if period not in self.result_sets:
            self.result_sets[period] = RatioResultSet(period)
        self.result_sets[period].add_result(result)

    def get_context_groups(self) -> Dict[str, List[RatioResult]]:
        """Group ratio results by reporting period"""
        log_debug("\nGrouping ratio results by context period")
        
        groups = defaultdict(list)
        for bs_context, result_set in self.result_sets.items():
            period = self.data_manager.get_period_for_context(bs_context, 'BS')
            if period[0] is None:
                period_str = f"{period[1].strftime('%Y-%m-%d')}"
            else:
                period_str = f"{period[0].strftime('%Y-%m-%d')} - {period[1].strftime('%Y-%m-%d')}"
                
            for result in result_set.results.values():
                groups[period_str].append(result)
                
        # Sort by end date
        def get_end_date(period_str):
            if ' - ' in period_str:
                return datetime.strptime(period_str.split(' - ')[1], '%Y-%m-%d')
            return datetime.strptime(period_str, '%Y-%m-%d')
                
        sorted_groups = dict(sorted(groups.items(), key=lambda x: get_end_date(x[0])))
        return sorted_groups

    def _assess_data_quality(self, result: RatioResult) -> str:
        """Assess data quality of ratio result"""
        try:
            # quality_factors = [] # unused!
            
            # Check if all required contexts are present
            expected_contexts = {'BS', 'IS'} if result.ratio_name in ['Return on Assets'] else {'BS'} if result.ratio_name in ['Current Ratio', 'Debt to Equity Ratio'] else {'IS'}
            has_all_contexts = all(ctx in result.context_ids for ctx in expected_contexts)
            
            # Check input quality
            all_inputs_valid = all(value is not None for _, (value, _, _) in result.inputs_used.items())
            has_custom_inputs = any(is_custom for _, (_, _, is_custom) in result.inputs_used.items())
            
            if has_all_contexts and all_inputs_valid and not has_custom_inputs:
                return "High"
            elif has_all_contexts and all_inputs_valid:
                return "Medium"
            else:
                return "Low"
                
        except Exception as e:
            log_debug(f"Error assessing data quality: {str(e)}")
            return "Unknown"

    def get_ratio_trends(self) -> Dict[str, List[Tuple[str, float]]]:
        """Get trends for all ratios across periods"""
        log_debug("\n=== Analyzing Ratio Trends ===")
        
        trends = defaultdict(list)
        context_groups = self.get_context_groups()
        
        for period_str, results in context_groups.items():
            for result in results:
                trends[result.ratio_name].append((period_str, result.value))
                
        # Calculate period-over-period changes
        for ratio_name, values in trends.items():
            if len(values) > 1:
                log_debug(f"\nTrend for {ratio_name}:")
                for i in range(1, len(values)):
                    prev_value = values[i-1][1]
                    curr_value = values[i][1]
                    if prev_value != 0:
                        change_pct = ((curr_value - prev_value) / abs(prev_value)) * 100
                        log_debug(f"{values[i][0]}: {change_pct:+.2f}%")
        
        return dict(trends)

    def validate_result_set(self, result_set: RatioResultSet) -> bool:
        """Validate a result set before adding to analysis"""
        log_debug(f"\nValidating result set for context: {result_set.bs_context}")
        
        try:
            # Check if we already have results for this context
            if result_set.bs_context in self.result_sets:
                log_debug("Warning: Overwriting existing results for context")
                
            # Validate all results in set
            invalid_results = []
            for ratio_name, result in result_set.results.items():
                is_valid, errors = result._validate()
                if not is_valid:
                    invalid_results.append((ratio_name, errors))
                    
            if invalid_results:
                log_debug("Invalid results found:")
                for ratio_name, errors in invalid_results:
                    log_debug(f"- {ratio_name}: {', '.join(errors)}")
                return False
                
            return True
            
        except Exception as e:
            log_debug(f"Error validating result set: {str(e)}")
            return False


class FinancialRatioCalculator:
    """
    Calculates financial ratios using FinancialDataManager for value retrieval
    and CalculationTracker for comprehensive tracking.
    """
    
    def __init__(self, model_xbrl):
        self.model_xbrl = model_xbrl
        self.fact_provider = model_xbrl._fact_provider
        self.data_manager = self.fact_provider.financial_data_manager
        self.calculation_tracker = CalculationTracker()
        self.logger = logger
        self.input_tracker = RatioInputTracker(self.data_manager)
        self.ratio_analysis = RatioAnalysis(model_xbrl)
        log_debug("\nInitialized FinancialRatioCalculator with ratio analysis (FinancialRatioCalculator.__init__)")
        
    def _get_value(self, concept: str, context_id: str, statement_type: str) -> Tuple[Optional[float], Optional[ValueMetadata]]:
        """Enhanced value retrieval using contextRefs"""
        # log_debug(f"\nRetrieving value for {concept} using contextRef {context_id}")
        
        return self.data_manager.get_value_by_context(concept, context_id, statement_type)
        
    def calculate_ratio(self, ratio_name: str, period: tuple) -> List[RatioResult]:
        """Calculate a ratio and track inputs"""
        result = None
        try:
            if ratio_name == "Current Ratio":
                result = self.calculate_current_ratio(period)[0]
            # ... other ratio cases ...
            
            if result:
                self.input_tracker.track_ratio_calculation(
                    ratio_name=ratio_name,
                    period=period,
                    inputs_used=result.inputs_used,
                    component_details=result.component_details
                )
                
            return result
            
        except Exception as e:
            self.logger.log(f"Error calculating {ratio_name}: {str(e)}", include_trace=True)
            return None

    def get_inputs_data(self, context_pairs: List[Tuple[str, str]]) -> dict:
        """Get organized ratio inputs data"""
        log_debug("\n=== Getting Ratio Inputs Data ===")
        log_debug(f"Processing {len(context_pairs)} context pairs")
        
        headers = ['Input', 'Origin', 'QName']
        
        # Log context pairs
        for bs_ctx, is_ctx in context_pairs:
            bs_period = self.data_manager.get_period_for_context(bs_ctx, 'BS')
            if bs_period and bs_period[1]:
                header = f"As of {bs_period[1].strftime('%Y-%m-%d')}"
                headers.append(header)
                log_debug(f"Added header: {header}")
        
        # Get organized inputs
        log_debug("\nGetting organized inputs from input tracker")
        data = self.input_tracker.get_organized_inputs(context_pairs)
        log_debug(f"Got {len(data)} input rows")
        
        result = {
            'headers': headers,
            'data': data
        }
        
        log_debug("\nFinal inputs data structure:")
        log_debug(f"Headers: {headers}")
        log_debug(f"Data rows: {len(data)}")
        
        return result

    def prepare_ratio_data(self) -> dict:
        """Prepare ratio data with context-based grouping"""
        log_debug("\n=== Preparing Ratio Data (FinancialRatioCalculator.prepare_ratio_data) ===")
        
        try:
            # Get valid context pairs
            context_pairs = self.data_manager.find_matching_contexts()
            if not context_pairs:
                log_debug("No valid context pairs found (FinancialRatioCalculator.prepare_ratio_data)")
                return {'headers': [], 'data': []}
                    
            # Sort pairs by end date
            def get_end_date(ctx_pair):
                bs_ctx, is_ctx = ctx_pair
                bs_period = self.data_manager.fact_provider.filing_analyzer.statement_periods['BS'].get(bs_ctx)
                if bs_period and bs_period[1]:
                    return bs_period[1]
                is_period = self.data_manager.fact_provider.filing_analyzer.statement_periods['IS'].get(is_ctx)
                return is_period[1] if is_period else datetime.min
                        
            sorted_pairs = sorted(context_pairs, key=get_end_date)
            
            # Calculate ratios for each context set
            for bs_ctx, is_ctx in sorted_pairs:
                log_debug(f"\nProcessing context set for BS: {bs_ctx}, IS: {is_ctx}")
                
                # Get CF context BEFORE calculating ratios
                cf_contexts = self.data_manager.find_matching_cf_contexts(bs_ctx)
                cf_ctx = cf_contexts[0] if cf_contexts else None
                log_debug(f"Found CF context: {cf_ctx}")

                if cf_contexts:
                    cf_period = self.data_manager.get_period_for_context(cf_ctx, 'CF')
                    if cf_period:
                        log_debug(f"CF Period: {cf_period[0].strftime('%Y-%m-%d') if cf_period[0] else 'None'} to {cf_period[1].strftime('%Y-%m-%d')}")
                
                result_set = self.calculate_ratios_for_contexts(bs_ctx, is_ctx, cf_ctx)
                if result_set and result_set.results:
                    log_debug(f"Calculated {len(result_set.results)} ratios")
                    self.ratio_analysis.add_result_set(result_set)
                    
            # Get formatted data from ratio analysis
            ratio_data = self.ratio_analysis.prepare_excel_data()
            log_debug(f"Prepared {len(ratio_data.get('data', [])) if ratio_data else 0} ratio records")
            
            return ratio_data
            
        except Exception as e:
            log_debug(f"Error preparing ratio data: {str(e)}", include_trace=True)
            return {'headers': [], 'data': []}

    def _add_ratio_results(self, result_set: RatioResultSet, results: List[RatioResult]):
        """Helper to add ratio results to set with logging"""
        for result in results:
            if result:
                result_set.add_result(result)
                log_debug(f"Added {result.ratio_name}: {result.value:.4f}")
                           
    def calculate_current_ratio(self, bs_context: str) -> List[RatioResult]:
        """Calculate current ratio using data manager"""
        log_debug(f"\n=== Calculating Current Ratio ===")
        log_debug(f"Using BS context: {bs_context}")
        
        self.calculation_tracker.start_calculation("ratio", "Current Ratio", bs_context)
        results = []
        
        try:
            # Get values through data manager
            current_assets, ca_metadata = self.data_manager.get_value_by_context(
                'current_assets', bs_context, 'BS')
            log_debug(f"Retrieved current assets: {current_assets}")
            
            current_liabs, cl_metadata = self.data_manager.get_value_by_context(
                'current_liabilities', bs_context, 'BS')
            log_debug(f"Retrieved current liabilities: {current_liabs}")
            
            if all(v is not None for v in [current_assets, current_liabs]) and current_liabs != 0:
                ratio = current_assets / current_liabs
                log_debug(f"Calculated ratio: {ratio:.4f}")
                
                # Create result with proper context mapping
                result = RatioResult(
                    ratio_name="Current Ratio",
                    value=ratio,
                    context_ids={'BS': bs_context},  # Using context_ids instead of period
                    inputs_used={
                        'current_assets': (current_assets, ca_metadata.qname, ca_metadata.is_custom),
                        'current_liabilities': (current_liabs, cl_metadata.qname, cl_metadata.is_custom)
                    }
                )
                
                results.append(result)
                self.calculation_tracker.complete_calculation(ratio)
                
            else:
                log_debug("Could not calculate - missing values or zero denominator")
                
            return results
            
        except Exception as e:
            log_debug(f"Error calculating current ratio: {str(e)}", include_trace=True)
            self.calculation_tracker.complete_calculation(None, "error", str(e))
            return results
    
    def calculate_debt_equity_ratio(self, bs_context: str) -> List[RatioResult]:
        """
        Calculate Debt/Equity ratio using context-based approach with 
        component tracking for total debt calculation
        """
        log_debug(f"\n=== Calculating Debt/Equity Ratio ===")
        log_debug(f"Using BS context: {bs_context}")
        
        self.calculation_tracker.start_calculation("ratio", "Debt to Equity", bs_context)
        results = []
        
        try:
            # Get total debt (may be compound value)
            total_debt, td_metadata = self.data_manager.get_value_by_context(
                'total_debt', bs_context, 'BS')
            log_debug(f"Retrieved total debt: {total_debt}")
            # # Excessive logging start (may be removed later)
            # if td_metadata:
            #     log_debug(f"Total debt components: {td_metadata.components if hasattr(td_metadata, 'components') else 'None'}")
            # # Excessive logging end
            
            # Get total equity
            total_equity, te_metadata = self.data_manager.get_value_by_context(
                'total_equity', bs_context, 'BS')
            log_debug(f"Retrieved total equity: {total_equity}")
            
            if all(v is not None for v in [total_debt, total_equity]) and total_equity != 0:
                ratio = total_debt / total_equity
                log_debug(f"Calculated D/E ratio: {ratio:.4f}")
                
                # Get component details if total_debt is a compound value
                component_details = None
                if hasattr(td_metadata, 'components') and td_metadata.components:
                    component_details = {
                        'total_debt': {
                            'value': total_debt,
                            'metadata': td_metadata,
                            'components': td_metadata.components
                        }
                    }
                    log_debug("Retrieved debt components")
                
                result = RatioResult(
                    ratio_name="Debt to Equity Ratio",
                    value=ratio,
                    context_ids={'BS': bs_context},
                    inputs_used={
                        'total_debt': (total_debt, td_metadata.qname, td_metadata.is_custom),
                        'total_equity': (total_equity, te_metadata.qname, te_metadata.is_custom)
                    },
                    component_details=component_details
                )
                
                # Track calculation - removed extra component_details argument
                self.calculation_tracker.track_ratio_calculation(
                    "Debt to Equity Ratio",
                    ratio,
                    {'BS': bs_context},
                    result.inputs_used
                )
                
                results.append(result)
                self.calculation_tracker.complete_calculation(ratio)
                
            else:
                log_debug("Could not calculate - missing values or zero equity")
                
            return results
            
        except Exception as e:
            log_debug(f"Error calculating D/E ratio: {str(e)}", include_trace=True)
            self.calculation_tracker.complete_calculation(None, "error", str(e))
            return results
            
    def calculate_gross_margin(self, is_context: str) -> List[RatioResult]:
        """Calculate gross margin using data manager"""
        log_debug(f"\n=== Calculating Gross Margin ===")
        log_debug(f"Using IS context: {is_context}")
        
        self.calculation_tracker.start_calculation("ratio", "Gross Margin", is_context)
        results = []
        
        try:
            # Get values through data manager
            revenue, rev_metadata = self.data_manager.get_value_by_context(
                'revenue', is_context, 'IS')
            log_debug(f"Retrieved revenue: {revenue}")
            
            cost_of_sales, cos_metadata = self.data_manager.get_value_by_context(
                'cost_of_sales', is_context, 'IS')
            log_debug(f"Retrieved cost of sales: {cost_of_sales}")
            
            if all(v is not None for v in [revenue, cost_of_sales]) and revenue != 0:
                gross_profit = revenue - cost_of_sales
                ratio = (gross_profit / revenue) 
                log_debug(f"Calculated gross margin: {ratio:.4f}")
                
                result = RatioResult(
                    ratio_name="Gross Margin",
                    value=ratio,
                    context_ids={'IS': is_context},  # Use context_ids instead of period
                    inputs_used={
                        'revenue': (revenue, rev_metadata.qname, rev_metadata.is_custom),
                        'cost_of_sales': (cost_of_sales, cos_metadata.qname, cos_metadata.is_custom)
                    }
                )
                
                results.append(result)
                self.calculation_tracker.complete_calculation(ratio)
                
            else:
                log_debug("Could not calculate - missing values or zero revenue")
                
            return results
            
        except Exception as e:
            log_debug(f"Error calculating gross margin: {str(e)}", include_trace=True)
            self.calculation_tracker.complete_calculation(None, "error", str(e))
            return results

    def calculate_net_margin(self, is_context: str) -> List[RatioResult]:
        """
        Calculate net margin using data manager and context-based approach
        Net Margin = Net Income / Revenue
        """
        log_debug(f"\n=== Calculating Net Margin ===")
        log_debug(f"Using IS context: {is_context}")
        
        self.calculation_tracker.start_calculation("ratio", "Net Margin", is_context)
        results = []
        
        try:
            # Get values through data manager using context
            net_income, ni_metadata = self.data_manager.get_value_by_context(
                'net_income', is_context, 'IS')
            log_debug(f"Retrieved net income: {net_income}")
            
            revenue, rev_metadata = self.data_manager.get_value_by_context(
                'revenue', is_context, 'IS')
            log_debug(f"Retrieved revenue: {revenue}")
            
            if all(v is not None for v in [net_income, revenue]) and revenue != 0:
                ratio = (net_income / revenue) 
                log_debug(f"Calculated net margin: {ratio:.4f}")
                
                result = RatioResult(
                    ratio_name="Net Margin",
                    value=ratio,
                    context_ids={'IS': is_context},  # Removed period parameter
                    inputs_used={
                        'net_income': (net_income, ni_metadata.qname, ni_metadata.is_custom),
                        'revenue': (revenue, rev_metadata.qname, rev_metadata.is_custom)
                    }
                )
                
                results.append(result)
                self.calculation_tracker.complete_calculation(ratio)
                
            else:
                log_debug("Could not calculate - missing values or zero revenue")
                
            return results
            
        except Exception as e:
            log_debug(f"Error calculating net margin: {str(e)}", include_trace=True)
            self.calculation_tracker.complete_calculation(None, "error", str(e))
            return results       
    
    def calculate_cost_of_sales_ratio(self, is_context: str) -> List[RatioResult]:
        """
        Calculate cost of sales ratio using context-based approach
        Cost of Sales Ratio = Cost of Sales / Revenue
        """
        log_debug(f"\n=== Calculating Cost of Sales Ratio ===")
        log_debug(f"Using IS context: {is_context}")
        
        self.calculation_tracker.start_calculation("ratio", "Cost of Sales Ratio", is_context)
        results = []
        
        try:
            # Get values through data manager using context
            cost_of_sales, cos_metadata = self.data_manager.get_value_by_context(
                'cost_of_sales', is_context, 'IS')
            log_debug(f"Retrieved cost of sales: {cost_of_sales}")
            
            revenue, rev_metadata = self.data_manager.get_value_by_context(
                'revenue', is_context, 'IS')
            log_debug(f"Retrieved revenue: {revenue}")
            
            if all(v is not None for v in [cost_of_sales, revenue]) and revenue != 0:
                ratio = (cost_of_sales / revenue) 
                log_debug(f"Calculated cost of sales ratio: {ratio:.4f}")
                
                result = RatioResult(
                    ratio_name="Cost of Sales Ratio",
                    value=ratio,
                    context_ids={'IS': is_context},
                    inputs_used={
                        'cost_of_sales': (cost_of_sales, cos_metadata.qname, cos_metadata.is_custom),
                        'revenue': (revenue, rev_metadata.qname, rev_metadata.is_custom)
                    }
                )
                
                results.append(result)
                self.calculation_tracker.complete_calculation(ratio)
                
            else:
                log_debug("Could not calculate - missing values or zero revenue")
                
            return results
            
        except Exception as e:
            log_debug(f"Error calculating cost of sales ratio: {str(e)}", include_trace=True)
            self.calculation_tracker.complete_calculation(None, "error", str(e))
            return results
        
    def calculate_return_on_assets(self, bs_context: str, is_context: str) -> List[RatioResult]:
        """Calculate ROA using context-based approach"""
        log_debug(f"\n=== Calculating Return on Assets ===")
        log_debug(f"Using contexts - BS: {bs_context}, IS: {is_context}")
        
        self.calculation_tracker.start_calculation("ratio", "Return on Assets", (bs_context, is_context))
        results = []
        
        try:
            # Get net income from IS using context
            net_income, ni_metadata = self.data_manager.get_value_by_context(
                'net_income', is_context, 'IS')
            log_debug(f"Retrieved net income: {net_income}")
            
            # Get total assets from BS using context
            total_assets, ta_metadata = self.data_manager.get_value_by_context(
                'total_assets', bs_context, 'BS')
            log_debug(f"Retrieved total assets: {total_assets}")
            
            if all(v is not None for v in [net_income, total_assets]) and total_assets != 0:
                ratio = net_income / total_assets 
                log_debug(f"Calculated ROA: {ratio:.4f}")
                
                result = RatioResult(
                    ratio_name="Return on Assets",
                    value=ratio,
                    context_ids={
                        'BS': bs_context,
                        'IS': is_context
                    },
                    inputs_used={
                        'net_income': (net_income, ni_metadata.qname, ni_metadata.is_custom),
                        'total_assets': (total_assets, ta_metadata.qname, ta_metadata.is_custom)
                    }
                )
                
                results.append(result)
                self.calculation_tracker.complete_calculation(ratio)
                
            else:
                log_debug("Could not calculate - missing values or zero assets")
                
            return results
            
        except Exception as e:
            log_debug(f"Error calculating ROA: {str(e)}", include_trace=True)
            self.calculation_tracker.complete_calculation(None, "error", str(e))
            return results

    def calculate_quick_ratio(self, bs_context: str) -> List[RatioResult]:
        """Calculate quick ratio using context-based approach"""
        log_debug(f"\n=== Calculating Quick Ratio ===")
        log_debug(f"Using BS context: {bs_context}")
        
        self.calculation_tracker.start_calculation("ratio", "Quick Ratio", bs_context)
        results = []
        
        try:
            # Get quick assets (compound value) through data manager
            quick_assets, qa_metadata = self.data_manager.get_value_by_context(
                'quick_assets', bs_context, 'BS')
            log_debug(f"Retrieved quick assets: {quick_assets}")
            
            # Track the main quick_assets value
            if quick_assets is not None:
                self.input_tracker.add_input(
                    name='quick_assets',
                    qname=qa_metadata.qname if qa_metadata else 'Calculated',
                    value=quick_assets,
                    context_id=bs_context,
                    statement='BS',
                    indent_level=0
                )
                
                # Track components if available
                if hasattr(qa_metadata, 'components'):
                    log_debug(f"Processing components: {qa_metadata.components}")
                    for component_name, component_value in qa_metadata.components.items():
                        # Get component metadata through data manager
                        comp_value, comp_metadata = self.data_manager.get_value_by_context(
                            component_name, bs_context, 'BS')
                        
                        self.input_tracker.add_input(
                            name=component_name,
                            qname=comp_metadata.qname if comp_metadata else 'Unknown',
                            value=component_value,
                            context_id=bs_context,
                            statement='BS',
                            indent_level=1,
                            calculation_rule="subtract"
                        )
                        log_debug(f"Added component {component_name}: {component_value}")
            
            # Get current liabilities
            current_liabs, cl_metadata = self.data_manager.get_value_by_context(
                'current_liabilities', bs_context, 'BS')
            log_debug(f"Retrieved current liabilities: {current_liabs}")
            
            if all(v is not None for v in [quick_assets, current_liabs]) and current_liabs != 0:
                ratio = quick_assets / current_liabs
                log_debug(f"Calculated quick ratio: {ratio:.4f}")
                
                result = RatioResult(
                    ratio_name="Quick Ratio",
                    value=ratio,
                    context_ids={'BS': bs_context},
                    inputs_used={
                        'quick_assets': (quick_assets, qa_metadata.qname, qa_metadata.is_custom),
                        'current_liabilities': (current_liabs, cl_metadata.qname, cl_metadata.is_custom)
                    },
                    component_details=qa_metadata.components if hasattr(qa_metadata, 'components') else None
                )
                
                results.append(result)
                self.calculation_tracker.complete_calculation(ratio)
            else:
                log_debug("Could not calculate - missing values or zero denominator")
                
            return results
            
        except Exception as e:
            log_debug(f"Error calculating quick ratio: {str(e)}", include_trace=True)
            self.calculation_tracker.complete_calculation(None, "error", str(e))
            return results
        
    def calculate_equity_multiplier(self, bs_context: str) -> List[RatioResult]:
        """
        Calculate equity multiplier using context-based approach
        Equity Multiplier = Total Assets / Total Equity
        """
        log_debug(f"\n=== Calculating Equity Multiplier ===")
        log_debug(f"Using BS context: {bs_context}")
        
        self.calculation_tracker.start_calculation("ratio", "Equity Multiplier", bs_context)
        results = []
        
        try:
            # Get values through data manager
            total_assets, ta_metadata = self.data_manager.get_value_by_context(
                'total_assets', bs_context, 'BS')
            log_debug(f"Retrieved total assets: {total_assets}")
            
            total_equity, te_metadata = self.data_manager.get_value_by_context(
                'total_equity', bs_context, 'BS')
            log_debug(f"Retrieved total equity: {total_equity}")
            
            if all(v is not None for v in [total_assets, total_equity]) and total_equity != 0:
                ratio = total_assets / total_equity
                log_debug(f"Calculated ratio: {ratio:.4f}")
                
                # Optional calculation for interpretation
                debt_portion = (ratio - 1) / ratio if ratio != 0 else 0
                log_debug(f"Assets financed by debt: {debt_portion:.2%}")
                
                result = RatioResult(
                    ratio_name="Equity Multiplier",
                    value=ratio,
                    context_ids={'BS': bs_context},
                    inputs_used={
                        'total_assets': (total_assets, ta_metadata.qname, ta_metadata.is_custom),
                        'total_equity': (total_equity, te_metadata.qname, te_metadata.is_custom)
                    }
                )
                
                results.append(result)
                self.calculation_tracker.complete_calculation(ratio)
                
            else:
                log_debug("Could not calculate - missing values or zero equity")
                
            return results
            
        except Exception as e:
            log_debug(f"Error calculating equity multiplier: {str(e)}", include_trace=True)
            self.calculation_tracker.complete_calculation(None, "error", str(e))
            return results
        
    def calculate_debt_to_capital_ratio(self, bs_context: str) -> List[RatioResult]:
        """
        Calculate debt to capital ratio using context-based approach
        Debt to Capital = Total Debt / Total Capital
        """
        log_debug(f"\n=== Calculating Debt to Capital Ratio ===")
        log_debug(f"Using BS context: {bs_context}")
        
        self.calculation_tracker.start_calculation("ratio", "Debt to Capital", bs_context)
        results = []
        
        try:
            # Get total debt and capital through data manager
            total_debt, td_metadata = self.data_manager.get_value_by_context(
                'total_debt', bs_context, 'BS')
            log_debug(f"Retrieved total debt: {total_debt}")
            
            total_capital, tc_metadata = self.data_manager.get_value_by_context(
                'total_capital', bs_context, 'BS')
            log_debug(f"Retrieved total capital: {total_capital}")
            
            if all(v is not None for v in [total_debt, total_capital]) and total_capital != 0:
                ratio = total_debt / total_capital
                log_debug(f"Calculated ratio: {ratio:.4f}")
                
                result = RatioResult(
                    ratio_name="Debt to Capital Ratio",
                    value=ratio,
                    context_ids={'BS': bs_context},
                    inputs_used={
                        'total_debt': (total_debt, td_metadata.qname, td_metadata.is_custom),
                        'total_capital': (total_capital, tc_metadata.qname, tc_metadata.is_custom)
                    }
                )
                
                results.append(result)
                self.calculation_tracker.complete_calculation(ratio)
                
            else:
                # Try alternative calculation using equity
                total_equity, te_metadata = self.data_manager.get_value_by_context(
                    'total_equity', bs_context, 'BS')
                log_debug(f"Retrieved total equity: {total_equity}")
                
                if all(v is not None for v in [total_debt, total_equity]):
                    calculated_capital = total_debt + total_equity
                    if calculated_capital != 0:
                        ratio = total_debt / calculated_capital
                        log_debug(f"Calculated alternative ratio: {ratio:.4f}")
                        
                        result = RatioResult(
                            ratio_name="Debt to Capital Ratio",
                            value=ratio,
                            context_ids={'BS': bs_context},
                            inputs_used={
                                'total_debt': (total_debt, td_metadata.qname, td_metadata.is_custom),
                                'total_equity': (total_equity, te_metadata.qname, te_metadata.is_custom)
                            },
                            calculation_variant="components"
                        )
                        
                        results.append(result)
                        self.calculation_tracker.complete_calculation(ratio)
                    else:
                        log_debug("Could not calculate - zero capital")
                else:
                    log_debug("Could not calculate - missing values")
                    
            return results
            
        except Exception as e:
            log_debug(f"Error calculating debt to capital ratio: {str(e)}", include_trace=True)
            self.calculation_tracker.complete_calculation(None, "error", str(e))
            return results        

    def calculate_return_on_equity(self, bs_context: str, is_context: str) -> List[RatioResult]:
        """Calculate ROE using context-based approach"""
        log_debug(f"\n=== Calculating Return on Equity ===")
        log_debug(f"Using contexts - BS: {bs_context}, IS: {is_context}")
        
        self.calculation_tracker.start_calculation("ratio", "Return on Equity", (bs_context, is_context))
        results = []
        
        try:
            # Get net income from IS using context
            net_income, ni_metadata = self.data_manager.get_value_by_context(
                'net_income', is_context, 'IS')
            log_debug(f"Retrieved net income: {net_income}")
            
            # Get total equity from BS using context
            total_equity, te_metadata = self.data_manager.get_value_by_context(
                'total_equity', bs_context, 'BS')
            log_debug(f"Retrieved total equity: {total_equity}")
            
            if all(v is not None for v in [net_income, total_equity]) and total_equity != 0:
                ratio = net_income / total_equity
                log_debug(f"Calculated ROE: {ratio:.4f}")
                
                result = RatioResult(
                    ratio_name="Return on Equity",
                    value=ratio,
                    context_ids={
                        'BS': bs_context,
                        'IS': is_context
                    },
                    inputs_used={
                        'net_income': (net_income, ni_metadata.qname, ni_metadata.is_custom),
                        'total_equity': (total_equity, te_metadata.qname, te_metadata.is_custom)
                    }
                )
                
                results.append(result)
                self.calculation_tracker.complete_calculation(ratio)
                
            else:
                log_debug("Could not calculate - missing values or zero equity")
                
            return results
            
        except Exception as e:
            log_debug(f"Error calculating ROE: {str(e)}", include_trace=True)
            self.calculation_tracker.complete_calculation(None, "error", str(e))
            return results
        
    def calculate_cash_ratio(self, bs_context: str) -> List[RatioResult]:
        """Calculate cash ratio using context-based approach"""
        log_debug(f"\n=== Calculating Cash Ratio ===")
        log_debug(f"Using BS context: {bs_context}")
        
        self.calculation_tracker.start_calculation("ratio", "Cash Ratio", bs_context)
        results = []
        
        try:
            # Try direct cash and equivalents first
            cash_and_equiv, cae_metadata = self.data_manager.get_value_by_context(
                'cash_and_equivalents', bs_context, 'BS')
            
            if cash_and_equiv is None:
                # Try getting individual components
                log_debug("Direct value not found, trying component calculation")
                cash, cash_metadata = self.data_manager.get_value_by_context(
                    'cash', bs_context, 'BS')
                log_debug(f"Retrieved cash: {cash}")
                
                cash_equiv, ce_metadata = self.data_manager.get_value_by_context(
                    'cash_equivalents', bs_context, 'BS')
                log_debug(f"Retrieved cash equivalents: {cash_equiv}")
                
                if cash is not None and cash_equiv is not None:
                    cash_and_equiv = cash + cash_equiv
                    log_debug(f"Calculated total cash and equivalents: {cash_and_equiv}")
            else:
                log_debug(f"Found direct cash and equivalents value: {cash_and_equiv}")
            
            # Get current liabilities
            current_liabs, cl_metadata = self.data_manager.get_value_by_context(
                'current_liabilities', bs_context, 'BS')
            log_debug(f"Retrieved current liabilities: {current_liabs}")
            
            if cash_and_equiv is not None and current_liabs is not None and current_liabs != 0:
                ratio = cash_and_equiv / current_liabs
                log_debug(f"Calculated cash ratio: {ratio:.4f}")
                
                result = RatioResult(
                    ratio_name="Cash Ratio",
                    value=ratio,
                    context_ids={'BS': bs_context},
                    inputs_used={
                        'cash_and_equivalents': (cash_and_equiv, cae_metadata.qname if cae_metadata else 'Calculated', 
                                            cae_metadata.is_custom if cae_metadata else True),
                        'current_liabilities': (current_liabs, cl_metadata.qname, cl_metadata.is_custom)
                    }
                )
                
                results.append(result)
                self.calculation_tracker.complete_calculation(ratio)
                
            else:
                log_debug("Could not calculate - missing values or zero denominator")
                
            return results
            
        except Exception as e:
            log_debug(f"Error calculating cash ratio: {str(e)}", include_trace=True)
            self.calculation_tracker.complete_calculation(None, "error", str(e))
            return results
            
    def calculate_debt_to_assets(self, bs_context: str) -> List[RatioResult]:
        """Calculate debt-to-assets ratio using context-based approach"""
        log_debug(f"\n=== Calculating Debt to Assets Ratio ===")
        log_debug(f"Using BS context: {bs_context}")
        
        self.calculation_tracker.start_calculation("ratio", "Debt to Assets", bs_context)
        results = []
        
        try:
            # Get total debt through data manager
            total_debt, td_metadata = self.data_manager.get_value_by_context(
                'total_debt', bs_context, 'BS')
            log_debug(f"Retrieved total debt: {total_debt}")
            
            # Get total assets
            total_assets, ta_metadata = self.data_manager.get_value_by_context(
                'total_assets', bs_context, 'BS')
            log_debug(f"Retrieved total assets: {total_assets}")
            
            if all(v is not None for v in [total_debt, total_assets]) and total_assets != 0:
                ratio = total_debt / total_assets
                log_debug(f"Calculated debt-to-assets ratio: {ratio:.4f}")
                
                result = RatioResult(
                    ratio_name="Debt to Assets Ratio",
                    value=ratio,
                    context_ids={'BS': bs_context},
                    inputs_used={
                        'total_debt': (total_debt, td_metadata.qname, td_metadata.is_custom),
                        'total_assets': (total_assets, ta_metadata.qname, ta_metadata.is_custom)
                    }
                )
                
                results.append(result)
                self.calculation_tracker.complete_calculation(ratio)
                
            else:
                log_debug("Could not calculate - missing values or zero assets")
                
            return results
            
        except Exception as e:
            log_debug(f"Error calculating debt-to-assets ratio: {str(e)}", include_trace=True)
            self.calculation_tracker.complete_calculation(None, "error", str(e))
            return results   
        
    def calculate_equity_ratio(self, bs_context: str) -> List[RatioResult]:
        """Calculate equity ratio using context-based approach"""
        log_debug(f"\n=== Calculating Equity Ratio ===")
        log_debug(f"Using BS context: {bs_context}")
        
        self.calculation_tracker.start_calculation("ratio", "Equity Ratio", bs_context)
        results = []
        
        try:
            # Get total equity
            total_equity, te_metadata = self.data_manager.get_value_by_context(
                'total_equity', bs_context, 'BS')
            log_debug(f"Retrieved total equity: {total_equity}")
            
            # Get total assets
            total_assets, ta_metadata = self.data_manager.get_value_by_context(
                'total_assets', bs_context, 'BS')
            log_debug(f"Retrieved total assets: {total_assets}")
            
            if all(v is not None for v in [total_equity, total_assets]) and total_assets != 0:
                ratio = total_equity / total_assets
                log_debug(f"Calculated equity ratio: {ratio:.4f}")
                
                result = RatioResult(
                    ratio_name="Equity Ratio",
                    value=ratio,
                    context_ids={'BS': bs_context},
                    inputs_used={
                        'total_equity': (total_equity, te_metadata.qname, te_metadata.is_custom),
                        'total_assets': (total_assets, ta_metadata.qname, ta_metadata.is_custom)
                    }
                )
                
                results.append(result)
                self.calculation_tracker.complete_calculation(ratio)
                
            else:
                log_debug("Could not calculate - missing values or zero assets")
                
            return results
            
        except Exception as e:
            log_debug(f"Error calculating equity ratio: {str(e)}", include_trace=True)
            self.calculation_tracker.complete_calculation(None, "error", str(e))
            return results         
  
    def calculate_ebit_margin(self, is_context: str) -> List[RatioResult]:
        """Calculate EBIT margin using context-based approach"""
        log_debug(f"\n=== Calculating EBIT Margin ===")
        log_debug(f"Using IS context: {is_context}")
        
        self.calculation_tracker.start_calculation("ratio", "EBIT Margin", is_context)
        results = []
        
        try:
            # Get EBIT through data manager
            ebit, ebit_metadata = self.data_manager.get_value_by_context(
                'ebit', is_context, 'IS')
            log_debug(f"Retrieved EBIT: {ebit}")
            
            # Track EBIT value and its components
            if ebit is not None:
                self.input_tracker.add_input(
                    name='ebit',
                    qname=ebit_metadata.qname if ebit_metadata else 'Calculated',
                    value=ebit,
                    context_id=is_context,
                    statement='IS',
                    indent_level=0
                )
                
                # Track EBIT components if available
                if ebit_metadata and hasattr(ebit_metadata, 'components'):
                    for component_name, component_value in ebit_metadata.components.items():
                        comp_value, comp_metadata = self.data_manager.get_value_by_context(
                            component_name, is_context, 'IS')
                        if comp_value is not None:
                            # self.input_tracker.add_input(
                            #     name=component_name,
                            #     qname=comp_metadata.qname if comp_metadata else 'Unknown',
                            #     value=comp_value,
                            #     context_id=is_context,
                            #     statement='IS',
                            #     indent_level=1,
                            #     sign="+" if comp_value >= 0 else "-"
                            # )                            
                            self.input_tracker.add_input(
                                name=component_name,
                                qname=comp_metadata.qname if comp_metadata else 'Unknown',
                                value=comp_value,
                                context_id=is_context,
                                statement='IS',
                                indent_level=1,
                                calculation_rule="subtract"  # Changed from sign parameter
                            )
            
            revenue, rev_metadata = self.data_manager.get_value_by_context(
                'revenue', is_context, 'IS')
            log_debug(f"Retrieved revenue: {revenue}")
            
            # Track revenue
            if revenue is not None:
                self.input_tracker.add_input(
                    name='revenue',
                    qname=rev_metadata.qname if rev_metadata else 'Unknown',
                    value=revenue,
                    context_id=is_context,
                    statement='IS',
                    indent_level=0
                )
            
            if all(v is not None for v in [ebit, revenue]) and revenue != 0:
                ratio = ebit / revenue
                log_debug(f"Calculated EBIT margin: {ratio:.4f}")
                
                result = RatioResult(
                    ratio_name="EBIT Margin",
                    value=ratio,
                    context_ids={'IS': is_context},
                    inputs_used={
                        'ebit': (ebit, ebit_metadata.qname, ebit_metadata.is_custom),
                        'revenue': (revenue, rev_metadata.qname, rev_metadata.is_custom)
                    }
                )
                
                results.append(result)
                self.calculation_tracker.complete_calculation(ratio)
                
            else:
                log_debug("Could not calculate - missing values or zero revenue")
                
            return results
            
        except Exception as e:
            log_debug(f"Error calculating EBIT margin: {str(e)}", include_trace=True)
            self.calculation_tracker.complete_calculation(None, "error", str(e))
            return results
    
    def calculate_ebitda_margin(self, is_context: str) -> List[RatioResult]:
        """Calculate EBITDA margin using context-based approach"""
        log_debug(f"\n=== Calculating EBITDA Margin ===")
        log_debug(f"Using IS context: {is_context}")
        
        self.calculation_tracker.start_calculation("ratio", "EBITDA Margin", is_context)
        results = []
        
        try:
            # Get EBITDA through data manager
            ebitda, ebitda_metadata = self.data_manager.get_value_by_context(
                'ebitda', is_context, 'IS')
            log_debug(f"Retrieved EBITDA: {ebitda}")
            
            revenue, rev_metadata = self.data_manager.get_value_by_context(
                'revenue', is_context, 'IS')
            log_debug(f"Retrieved revenue: {revenue}")
            
            if all(v is not None for v in [ebitda, revenue]) and revenue != 0:
                ratio = ebitda / revenue
                log_debug(f"Calculated EBITDA margin: {ratio:.4f}")
                
                result = RatioResult(
                    ratio_name="EBITDA Margin",
                    value=ratio,
                    context_ids={'IS': is_context},
                    inputs_used={
                        'ebitda': (ebitda, ebitda_metadata.qname, ebitda_metadata.is_custom),
                        'revenue': (revenue, rev_metadata.qname, rev_metadata.is_custom)
                    }
                )
                
                results.append(result)
                self.calculation_tracker.complete_calculation(ratio)
                
            else:
                log_debug("Could not calculate - missing values or zero revenue")
                
            return results
            
        except Exception as e:
            log_debug(f"Error calculating EBITDA margin: {str(e)}", include_trace=True)
            self.calculation_tracker.complete_calculation(None, "error", str(e))
            return results

    def calculate_interest_coverage(self, is_context: str) -> List[RatioResult]:
        """Calculate interest coverage (EBIT/Interest) using context-based approach"""
        log_debug(f"\n=== Calculating Interest Coverage ===")
        log_debug(f"Using IS context: {is_context}")
        
        self.calculation_tracker.start_calculation("ratio", "Interest Coverage", is_context)
        results = []
        
        try:
            # Get values through data manager
            ebit, ebit_metadata = self.data_manager.get_value_by_context(
                'ebit', is_context, 'IS')
            log_debug(f"Retrieved EBIT: {ebit}")
            
            interest, int_metadata = self.data_manager.get_value_by_context(
                'interest_expense', is_context, 'IS')
            log_debug(f"Retrieved interest expense: {interest}")
            
            if all(v is not None for v in [ebit, interest]) and interest != 0:
                ratio = ebit / interest
                log_debug(f"Calculated interest coverage: {ratio:.4f}")
                
                result = RatioResult(
                    ratio_name="Interest Coverage Ratio",
                    value=ratio,
                    context_ids={'IS': is_context},
                    inputs_used={
                        'ebit': (ebit, ebit_metadata.qname, ebit_metadata.is_custom),
                        'interest_expense': (interest, int_metadata.qname, int_metadata.is_custom)
                    }
                )
                
                results.append(result)
                self.calculation_tracker.complete_calculation(ratio)
                
            else:
                log_debug("Could not calculate - missing values or zero interest")
                
            return results
            
        except Exception as e:
            log_debug(f"Error calculating interest coverage: {str(e)}", include_trace=True)
            self.calculation_tracker.complete_calculation(None, "error", str(e))
            return results
    
    # See EBIT MARGIN
    # def calculate_operating_profit_margin(self, is_context: str) -> List[RatioResult]:
        """Calculate operating profit margin using EBIT/Revenue"""
        log_debug(f"\n=== Calculating Operating Profit Margin ===")
        log_debug(f"Using IS context: {is_context}")
        
        self.calculation_tracker.start_calculation("ratio", "Operating Profit Margin", is_context)
        results = []
        
        try:
            # Get EBIT and revenue through data manager
            ebit, ebit_metadata = self.data_manager.get_value_by_context(
                'ebit', is_context, 'IS')
            log_debug(f"Retrieved EBIT: {ebit}")
            
            revenue, rev_metadata = self.data_manager.get_value_by_context(
                'revenue', is_context, 'IS')
            log_debug(f"Retrieved revenue: {revenue}")
            
            if all(v is not None for v in [ebit, revenue]) and revenue != 0:
                ratio = ebit / revenue
                log_debug(f"Calculated operating profit margin: {ratio:.4f}")
                
                result = RatioResult(
                    ratio_name="Operating Profit Margin",
                    value=ratio,
                    context_ids={'IS': is_context},
                    inputs_used={
                        'ebit': (ebit, ebit_metadata.qname, ebit_metadata.is_custom),
                        'revenue': (revenue, rev_metadata.qname, rev_metadata.is_custom)
                    }
                )
                
                results.append(result)
                self.calculation_tracker.complete_calculation(ratio)
                
            else:
                log_debug("Could not calculate - missing values or zero revenue")
                
            return results
            
        except Exception as e:
            log_debug(f"Error calculating operating profit margin: {str(e)}", include_trace=True)
            self.calculation_tracker.complete_calculation(None, "error", str(e))
            return results

    def calculate_roce(self, bs_context: str, is_context: str) -> List[RatioResult]:
        """Calculate Return on Capital Employed (ROCE)"""
        log_debug(f"\n=== Calculating ROCE ===")
        log_debug(f"Using contexts - BS: {bs_context}, IS: {is_context}")
        
        self.calculation_tracker.start_calculation("ratio", "ROCE", (bs_context, is_context))
        results = []
        
        try:
            # Get EBIT from IS
            ebit, ebit_metadata = self.data_manager.get_value_by_context(
                'ebit', is_context, 'IS')
            log_debug(f"Retrieved EBIT: {ebit}")
            
            # Get total assets and current liabilities from BS
            total_assets, ta_metadata = self.data_manager.get_value_by_context(
                'total_assets', bs_context, 'BS')
            log_debug(f"Retrieved total assets: {total_assets}")
            
            current_liabs, cl_metadata = self.data_manager.get_value_by_context(
                'current_liabilities', bs_context, 'BS')
            log_debug(f"Retrieved current liabilities: {current_liabs}")
            
            if all(v is not None for v in [ebit, total_assets, current_liabs]):
                capital_employed = total_assets - current_liabs
                if capital_employed != 0:
                    ratio = ebit / capital_employed
                    log_debug(f"Calculated ROCE: {ratio:.4f}")
                    
                    result = RatioResult(
                        ratio_name="Return on Capital Employed",
                        value=ratio,
                        context_ids={
                            'BS': bs_context,
                            'IS': is_context
                        },
                        inputs_used={
                            'ebit': (ebit, ebit_metadata.qname, ebit_metadata.is_custom),
                            'total_assets': (total_assets, ta_metadata.qname, ta_metadata.is_custom),
                            'current_liabilities': (current_liabs, cl_metadata.qname, cl_metadata.is_custom)
                        }
                    )
                    
                    results.append(result)
                    self.calculation_tracker.complete_calculation(ratio)
                else:
                    log_debug("Could not calculate - zero capital employed")
            else:
                log_debug("Could not calculate - missing values")
                
            return results
            
        except Exception as e:
            log_debug(f"Error calculating ROCE: {str(e)}", include_trace=True)
            self.calculation_tracker.complete_calculation(None, "error", str(e))
            return results

    def calculate_asset_turnover(self, bs_context: str, is_context: str) -> List[RatioResult]:
        """Calculate Asset Turnover ratio"""
        log_debug(f"\n=== Calculating Asset Turnover ===")
        log_debug(f"Using contexts - BS: {bs_context}, IS: {is_context}")
        
        self.calculation_tracker.start_calculation("ratio", "Asset Turnover", (bs_context, is_context))
        results = []
        
        try:
            # Get revenue from IS
            revenue, rev_metadata = self.data_manager.get_value_by_context(
                'revenue', is_context, 'IS')
            log_debug(f"Retrieved revenue: {revenue}")
            
            # Get total assets from BS
            total_assets, ta_metadata = self.data_manager.get_value_by_context(
                'total_assets', bs_context, 'BS')
            log_debug(f"Retrieved total assets: {total_assets}")
            
            if all(v is not None for v in [revenue, total_assets]) and total_assets != 0:
                ratio = revenue / total_assets
                log_debug(f"Calculated asset turnover: {ratio:.4f}")
                
                result = RatioResult(
                    ratio_name="Asset Turnover",
                    value=ratio,
                    context_ids={
                        'BS': bs_context,
                        'IS': is_context
                    },
                    inputs_used={
                        'revenue': (revenue, rev_metadata.qname, rev_metadata.is_custom),
                        'total_assets': (total_assets, ta_metadata.qname, ta_metadata.is_custom)
                    }
                )
                
                results.append(result)
                self.calculation_tracker.complete_calculation(ratio)
                
            else:
                log_debug("Could not calculate - missing values or zero assets")
                
            return results
            
        except Exception as e:
            log_debug(f"Error calculating asset turnover: {str(e)}", include_trace=True)
            self.calculation_tracker.complete_calculation(None, "error", str(e))
            return results    

    def calculate_operating_cash_flow_ratio(self, bs_context: str, cf_context: str) -> List[RatioResult]:
        """Calculate operating cash flow ratio using context-based approach"""
        log_debug(f"\n=== Calculating Operating Cash Flow Ratio ===")
        log_debug(f"Using contexts - BS: {bs_context}, CF: {cf_context}")
        
        if not cf_context:
            log_debug("No CF context available - skipping calculation")
            return []
        self.calculation_tracker.start_calculation("ratio", "Operating Cash Flow Ratio", (bs_context, cf_context))
        results = []
        
        try:
            # Get operating cash flow from CF using context
            operating_cf, ocf_metadata = self.data_manager.get_value_by_context(
                'operating_cashflow', cf_context, 'CF')
            log_debug(f"Retrieved operating cash flow: {operating_cf}")
            
            # Get current liabilities from BS using context
            current_liabs, cl_metadata = self.data_manager.get_value_by_context(
                'current_liabilities', bs_context, 'BS')
            log_debug(f"Retrieved current liabilities: {current_liabs}")
            
            if all(v is not None for v in [operating_cf, current_liabs]) and current_liabs != 0:
                ratio = operating_cf / current_liabs
                log_debug(f"Calculated operating cash flow ratio: {ratio:.4f}")
                
                result = RatioResult(
                    ratio_name="Operating Cash Flow Ratio",
                    value=ratio,
                    context_ids={
                        'BS': bs_context,
                        'CF': cf_context
                    },
                    inputs_used={
                        'operating_cashflow': (operating_cf, ocf_metadata.qname, ocf_metadata.is_custom),
                        'current_liabilities': (current_liabs, cl_metadata.qname, cl_metadata.is_custom)
                    }
                )
                
                results.append(result)
                self.calculation_tracker.complete_calculation(ratio)
                
            else:
                log_debug("Could not calculate - missing values or zero denominator")
                
            return results
            
        except Exception as e:
            log_debug(f"Error calculating operating cash flow ratio: {str(e)}", include_trace=True)
            self.calculation_tracker.complete_calculation(None, "error", str(e))
            return results

    def calculate_operating_cash_flow_margin(self, cf_context: str, is_context: str) -> List[RatioResult]:
        """Calculate operating cash flow margin using context-based approach"""
        log_debug(f"\n=== Calculating Operating Cash Flow Margin ===")
        log_debug(f"Using contexts - CF: {cf_context}, IS: {is_context}")
        
        if not cf_context or not is_context:
            log_debug("Missing required context - skipping calculation")
            return []

        self.calculation_tracker.start_calculation("ratio", "Operating Cash Flow Margin", (cf_context, is_context))
        results = []
        
        try:
            # Get operating cash flow from CF using context
            operating_cf, ocf_metadata = self.data_manager.get_value_by_context(
                'operating_cashflow', cf_context, 'CF')
            log_debug(f"Retrieved operating cash flow: {operating_cf}")
            
            # Get revenue from IS using context
            revenue, rev_metadata = self.data_manager.get_value_by_context(
                'revenue', is_context, 'IS')
            log_debug(f"Retrieved revenue: {revenue}")
            
            if all(v is not None for v in [operating_cf, revenue]) and revenue != 0:
                ratio = operating_cf / revenue
                log_debug(f"Calculated operating cash flow margin: {ratio:.4f}")
                
                result = RatioResult(
                    ratio_name="Operating Cash Flow Margin",
                    value=ratio,
                    context_ids={
                        'CF': cf_context,
                        'IS': is_context
                    },
                    inputs_used={
                        'operating_cashflow': (operating_cf, ocf_metadata.qname, ocf_metadata.is_custom),
                        'revenue': (revenue, rev_metadata.qname, rev_metadata.is_custom)
                    }
                )
                
                results.append(result)
                self.calculation_tracker.complete_calculation(ratio)
                
            else:
                log_debug("Could not calculate - missing values or zero revenue")
                
            return results
            
        except Exception as e:
            log_debug(f"Error calculating operating cash flow margin: {str(e)}", include_trace=True)
            self.calculation_tracker.complete_calculation(None, "error", str(e))
            return results

    def calculate_inventory_turnover(self, bs_context: str, is_context: str) -> List[RatioResult]:
        """Calculate inventory turnover using context-based approach"""
        log_debug(f"\n=== Calculating Inventory Turnover ===")
        log_debug(f"Using contexts - BS: {bs_context}, IS: {is_context}")
        
        self.calculation_tracker.start_calculation("ratio", "Inventory Turnover", (bs_context, is_context))
        results = []
        
        try:
            # Get cost of sales from IS using context
            cost_of_sales, cos_metadata = self.data_manager.get_value_by_context(
                'cost_of_sales', is_context, 'IS')
            log_debug(f"Retrieved cost of sales: {cost_of_sales}")
            
            # Get inventory from BS using context
            inventory, inv_metadata = self.data_manager.get_value_by_context(
                'inventory', bs_context, 'BS')
            log_debug(f"Retrieved inventory: {inventory}")
            
            if all(v is not None for v in [cost_of_sales, inventory]) and inventory != 0:
                ratio = cost_of_sales / inventory
                log_debug(f"Calculated inventory turnover: {ratio:.4f}")
                
                result = RatioResult(
                    ratio_name="Inventory Turnover",
                    value=ratio,
                    context_ids={
                        'BS': bs_context,
                        'IS': is_context
                    },
                    inputs_used={
                        'cost_of_sales': (cost_of_sales, cos_metadata.qname, cos_metadata.is_custom),
                        'inventory': (inventory, inv_metadata.qname, inv_metadata.is_custom)
                    }
                )
                
                results.append(result)
                self.calculation_tracker.complete_calculation(ratio)
                
            else:
                log_debug("Could not calculate - missing values or zero inventory")
                
            return results
            
        except Exception as e:
            log_debug(f"Error calculating inventory turnover: {str(e)}", include_trace=True)
            self.calculation_tracker.complete_calculation(None, "error", str(e))
            return results

    def calculate_dividend_payout(self, is_context: str) -> List[RatioResult]:
        """Calculate dividend payout ratio using context-based approach"""
        log_debug(f"\n=== Calculating Dividend Payout Ratio ===")
        log_debug(f"Using IS context: {is_context}")
        
        self.calculation_tracker.start_calculation("ratio", "Dividend Payout", is_context)
        results = []
        
        try:
            # Get dividends and net income using context
            dividends, div_metadata = self.data_manager.get_value_by_context(
                'dividends', is_context, 'IS')
            log_debug(f"Retrieved dividends: {dividends}")
            
            net_income, ni_metadata = self.data_manager.get_value_by_context(
                'net_income', is_context, 'IS')
            log_debug(f"Retrieved net income: {net_income}")
            
            if all(v is not None for v in [dividends, net_income]) and net_income != 0:
                ratio = dividends / net_income
                log_debug(f"Calculated dividend payout ratio: {ratio:.4f}")
                
                result = RatioResult(
                    ratio_name="Dividend Payout Ratio",
                    value=ratio,
                    context_ids={'IS': is_context},
                    inputs_used={
                        'dividends': (dividends, div_metadata.qname, div_metadata.is_custom),
                        'net_income': (net_income, ni_metadata.qname, ni_metadata.is_custom)
                    }
                )
                
                results.append(result)
                self.calculation_tracker.complete_calculation(ratio)
                
            else:
                log_debug("Could not calculate - missing values or zero net income")
                
            return results
            
        except Exception as e:
            log_debug(f"Error calculating dividend payout ratio: {str(e)}", include_trace=True)
            self.calculation_tracker.complete_calculation(None, "error", str(e))
            return results    

    def calculate_net_debt_to_ebitda(self, bs_context: str, is_context: str) -> List[RatioResult]:
        """Calculate Net Debt to EBITDA ratio using context-based approach"""
        log_debug(f"\n=== Calculating Net Debt to EBITDA ===")
        log_debug(f"Using contexts - BS: {bs_context}, IS: {is_context}")
        
        self.calculation_tracker.start_calculation("ratio", "Net Debt to EBITDA", (bs_context, is_context))
        results = []
        
        try:
            # Get net debt (compound value) through data manager
            net_debt, nd_metadata = self.data_manager.get_value_by_context(
                'net_debt', bs_context, 'BS')
            log_debug(f"Retrieved net debt: {net_debt}")
            
            # Get EBITDA
            ebitda, ebitda_metadata = self.data_manager.get_value_by_context(
                'ebitda', is_context, 'IS')
            log_debug(f"Retrieved EBITDA: {ebitda}")
            
            if all(v is not None for v in [net_debt, ebitda]) and ebitda != 0:
                ratio = net_debt / ebitda
                log_debug(f"Calculated Net Debt to EBITDA: {ratio:.4f}")
                
                result = RatioResult(
                    ratio_name="Net Debt to EBITDA",
                    value=ratio,
                    context_ids={
                        'BS': bs_context,
                        'IS': is_context
                    },
                    inputs_used={
                        'net_debt': (net_debt, nd_metadata.qname, nd_metadata.is_custom),
                        'ebitda': (ebitda, ebitda_metadata.qname, ebitda_metadata.is_custom)
                    },
                    component_details=nd_metadata.components if hasattr(nd_metadata, 'components') else None
                )
                
                results.append(result)
                self.calculation_tracker.complete_calculation(ratio)
                
            else:
                log_debug("Could not calculate - missing values or zero EBITDA")
                
            return results
            
        except Exception as e:
            log_debug(f"Error calculating Net Debt to EBITDA: {str(e)}", include_trace=True)
            self.calculation_tracker.complete_calculation(None, "error", str(e))
            return results    

    def calculate_total_debt_to_ebitda(self, bs_context: str, is_context: str) -> List[RatioResult]:
        """Calculate Total Debt to EBITDA ratio using context-based approach"""
        log_debug(f"\n=== Calculating Total Debt to EBITDA ===")
        log_debug(f"Using contexts - BS: {bs_context}, IS: {is_context}")
        
        self.calculation_tracker.start_calculation("ratio", "Total Debt to EBITDA", (bs_context, is_context))
        results = []
        
        try:
            # Get total debt through data manager
            total_debt, td_metadata = self.data_manager.get_value_by_context(
                'total_debt', bs_context, 'BS')
            log_debug(f"Retrieved total debt: {total_debt}")
            
            # Get EBITDA
            ebitda, ebitda_metadata = self.data_manager.get_value_by_context(
                'ebitda', is_context, 'IS')
            log_debug(f"Retrieved EBITDA: {ebitda}")
            
            if all(v is not None for v in [total_debt, ebitda]) and ebitda != 0:
                ratio = total_debt / ebitda
                log_debug(f"Calculated Total Debt to EBITDA: {ratio:.4f}")
                
                result = RatioResult(
                    ratio_name="Total Debt to EBITDA",
                    value=ratio,
                    context_ids={
                        'BS': bs_context,
                        'IS': is_context
                    },
                    inputs_used={
                        'total_debt': (total_debt, td_metadata.qname, td_metadata.is_custom),
                        'ebitda': (ebitda, ebitda_metadata.qname, ebitda_metadata.is_custom)
                    },
                    component_details=td_metadata.components if hasattr(td_metadata, 'components') else None
                )
                
                results.append(result)
                self.calculation_tracker.complete_calculation(ratio)
                
            else:
                log_debug("Could not calculate - missing values or zero EBITDA")
                
            return results
            
        except Exception as e:
            log_debug(f"Error calculating Total Debt to EBITDA: {str(e)}", include_trace=True)
            self.calculation_tracker.complete_calculation(None, "error", str(e))
            return results

    def calculate_net_debt_to_equity(self, bs_context: str) -> List[RatioResult]:
        """Calculate Net Debt to Equity ratio using context-based approach"""
        log_debug(f"\n=== Calculating Net Debt to Equity ===")
        log_debug(f"Using BS context: {bs_context}")
        
        self.calculation_tracker.start_calculation("ratio", "Net Debt to Equity", bs_context)
        results = []
        
        try:
            # Get net debt (compound value) through data manager
            net_debt, nd_metadata = self.data_manager.get_value_by_context(
                'net_debt', bs_context, 'BS')
            log_debug(f"Retrieved net debt: {net_debt}")
            
            # Get total equity
            total_equity, te_metadata = self.data_manager.get_value_by_context(
                'total_equity', bs_context, 'BS')
            log_debug(f"Retrieved total equity: {total_equity}")
            
            if all(v is not None for v in [net_debt, total_equity]) and total_equity != 0:
                ratio = net_debt / total_equity
                log_debug(f"Calculated Net Debt to Equity: {ratio:.4f}")
                
                result = RatioResult(
                    ratio_name="Net Debt to Equity",
                    value=ratio,
                    context_ids={'BS': bs_context},
                    inputs_used={
                        'net_debt': (net_debt, nd_metadata.qname, nd_metadata.is_custom),
                        'total_equity': (total_equity, te_metadata.qname, te_metadata.is_custom)
                    },
                    component_details=nd_metadata.components if hasattr(nd_metadata, 'components') else None
                )
                
                results.append(result)
                self.calculation_tracker.complete_calculation(ratio)
                
            else:
                log_debug("Could not calculate - missing values or zero equity")
                
            return results
            
        except Exception as e:
            log_debug(f"Error calculating Net Debt to Equity: {str(e)}", include_trace=True)
            self.calculation_tracker.complete_calculation(None, "error", str(e))
            return results

    def calculate_long_term_debt_ratio(self, bs_context: str) -> List[RatioResult]:
        """Calculate Long-term Debt ratio using context-based approach"""
        log_debug(f"\n=== Calculating Long-term Debt Ratio ===")
        log_debug(f"Using BS context: {bs_context}")
        
        self.calculation_tracker.start_calculation("ratio", "Long-term Debt Ratio", bs_context)
        results = []
        
        try:
            # Get long-term debt through data manager
            long_term_debt, ltd_metadata = self.data_manager.get_value_by_context(
                'long_term_borrowings', bs_context, 'BS')
            log_debug(f"Retrieved long-term debt: {long_term_debt}")
            
            # Get total assets
            total_assets, ta_metadata = self.data_manager.get_value_by_context(
                'total_assets', bs_context, 'BS')
            log_debug(f"Retrieved total assets: {total_assets}")
            
            if all(v is not None for v in [long_term_debt, total_assets]) and total_assets != 0:
                ratio = long_term_debt / total_assets
                log_debug(f"Calculated Long-term Debt Ratio: {ratio:.4f}")
                
                result = RatioResult(
                    ratio_name="Long-term Debt Ratio",
                    value=ratio,
                    context_ids={'BS': bs_context},
                    inputs_used={
                        'long_term_borrowings': (long_term_debt, ltd_metadata.qname, ltd_metadata.is_custom),
                        'total_assets': (total_assets, ta_metadata.qname, ta_metadata.is_custom)
                    }
                )
                
                results.append(result)
                self.calculation_tracker.complete_calculation(ratio)
                
            else:
                log_debug("Could not calculate - missing values or zero assets")
                
            return results
            
        except Exception as e:
            log_debug(f"Error calculating Long-term Debt Ratio: {str(e)}", include_trace=True)
            self.calculation_tracker.complete_calculation(None, "error", str(e))
            return results

    def calculate_net_debt_to_capital(self, bs_context: str) -> List[RatioResult]:
        """Calculate Net Debt to Capital ratio using context-based approach"""
        log_debug(f"\n=== Calculating Net Debt to Capital ===")
        log_debug(f"Using BS context: {bs_context}")
        
        self.calculation_tracker.start_calculation("ratio", "Net Debt to Capital", bs_context)
        results = []
        
        try:
            # Get net debt (compound value) through data manager
            net_debt, nd_metadata = self.data_manager.get_value_by_context(
                'net_debt', bs_context, 'BS')
            log_debug(f"Retrieved net debt: {net_debt}")
            
            # Get total debt and total equity for capital calculation
            total_debt, td_metadata = self.data_manager.get_value_by_context(
                'total_debt', bs_context, 'BS')
            total_equity, te_metadata = self.data_manager.get_value_by_context(
                'total_equity', bs_context, 'BS')
            
            # Calculate total capital
            if total_debt is not None and total_equity is not None:
                total_capital = total_debt + total_equity
                log_debug(f"Calculated total capital: {total_capital}")
                
                if net_debt is not None and total_capital != 0:
                    ratio = net_debt / total_capital
                    log_debug(f"Calculated Net Debt to Capital: {ratio:.4f}")
                    
                    result = RatioResult(
                        ratio_name="Net Debt to Capital",
                        value=ratio,
                        context_ids={'BS': bs_context},
                        inputs_used={
                            'net_debt': (net_debt, nd_metadata.qname, nd_metadata.is_custom),
                            'total_debt': (total_debt, td_metadata.qname, td_metadata.is_custom),
                            'total_equity': (total_equity, te_metadata.qname, te_metadata.is_custom)
                        },
                        component_details=nd_metadata.components if hasattr(nd_metadata, 'components') else None
                    )
                    
                    results.append(result)
                    self.calculation_tracker.complete_calculation(ratio)
                else:
                    log_debug("Could not calculate - missing net debt or zero capital")
            else:
                log_debug("Could not calculate total capital - missing debt or equity values")
                
            return results
            
        except Exception as e:
            log_debug(f"Error calculating Net Debt to Capital: {str(e)}", include_trace=True)
            self.calculation_tracker.complete_calculation(None, "error", str(e))
            return results
    
    # add more ratios

    def calculate_ratios_for_contexts(self, bs_context: str, is_context: str = None, cf_context: str = None) -> RatioResultSet:
        """Calculate ratios for a context pair with enhanced tracking"""
        log_debug(f"\n=== Calculating Ratios for Contexts ===")
        log_debug(f"Called from: {traceback.extract_stack()[-2][2]}")  # Shows calling function
        log_debug(f"BS Context: {bs_context}")
        log_debug(f"IS Context: {is_context}")
        log_debug(f"CF Context: {cf_context}")  
        
        # Excessive logging
        # if cf_context:
        #     log_debug(f"\nValidating CF context: {cf_context}")
        #     cf_period = self.data_manager.get_period_for_context(cf_context, 'CF')
        #     if cf_period:
        #         log_debug(f"CF Period validated: {cf_period[0].strftime('%Y-%m-%d') if cf_period[0] else 'None'} to {cf_period[1].strftime('%Y-%m-%d')}")        
        
        try:
            # Create result set using BS context
            result_set = RatioResultSet(bs_context, self.data_manager)
            
            # Start calculation tracking
            self.calculation_tracker.start_calculation(
                "ratio_set", 
                "period_ratios", 
                bs_context  # Use context as identifier instead of period
            )
            
            try:
                # Track base financial values first
                log_debug("\nTracking base financial values...")
                base_concepts = [
                    ('total_assets', 'BS'), ('total_liabilities', 'BS'), ('total_equity', 'BS'),
                    ('revenue', 'IS'), ('net_income', 'IS')
                ]
                for concept, stmt_type in base_concepts:
                    context_id = bs_context if stmt_type == 'BS' else is_context
                    value, metadata = self.data_manager.get_value_by_context(concept, context_id, stmt_type)
                    if value is not None:
                        log_debug(f"Adding base value {concept}: {value}")
                        self.input_tracker.add_input(
                            name=concept,
                            qname=metadata.qname if metadata else 'Unknown',
                            value=value,
                            context_id=context_id,
                            statement=stmt_type,
                            indent_level=0
                        )

                # Track compound values with components
                log_debug("\nTracking compound values...")
                compounds = ['total_debt', 'total_capital', 'working_capital', 'gross_profit','net_debt']
                for compound in compounds:
                    # First check if it's a registered compound
                    is_compound = compound in self.data_manager.component_registry._registry
                    if not is_compound:
                        continue  # Skip if not a true compound value
                        
                    value, metadata = self.data_manager.get_value_by_context(compound, bs_context, 'BS')
                    if value is not None:
                        log_debug(f"Found registered compound {compound}: {value}")
                        # Add main compound value
                        self.input_tracker.add_input(
                            name=compound,
                            qname=metadata.qname if metadata else 'Calculated',
                            value=value,
                            context_id=bs_context,
                            statement='BS',
                            indent_level=0
                        )
                        
                        # Get components from registry
                        components = self.data_manager.component_registry.get_components(compound)
                        if components:
                            log_debug(f"Processing {len(components)} registered components for {compound}")
                            for component in components:
                                comp_value, comp_metadata = self.data_manager.get_value_by_context(
                                    component, bs_context, 'BS')
                                if comp_value is not None:
                                    self.input_tracker.add_input(
                                        name=component,
                                        qname=comp_metadata.qname if comp_metadata else 'Unknown',
                                        value=comp_value,
                                        context_id=bs_context,
                                        statement='BS',
                                        indent_level=1,
                                        calculation_rule="subtract"
                                    )
                                    log_debug(f"Added registered component {component}: {comp_value}")
                #################################################                                    
                # Calculate BS ratios
                if bs_context:
                    try:
                        log_debug("\nCalculating BS ratios...")
                        
                        # Current Ratio
                        current_ratio_results = self.calculate_current_ratio(bs_context)
                        if current_ratio_results:
                            for result in current_ratio_results:
                                result_set.add_result(result)
                                log_debug(f"Added Current Ratio: {result.value:.4f}")
                                self.calculation_tracker.track_ratio_calculation(
                                    "Current Ratio",
                                    result.value,
                                    {'BS': bs_context},
                                    result.inputs_used
                                )                                
                                
                    except Exception as e:
                        log_debug(f"Error calculating Current Ratio: {str(e)}", include_trace=True)                    
                    
                    try:
                        # Debt/Equity Ratio
                        de_ratio_results = self.calculate_debt_equity_ratio(bs_context)
                        if de_ratio_results:
                            for result in de_ratio_results:
                                result_set.add_result(result)
                                log_debug(f"Added D/E Ratio: {result.value:.4f}")
                                self.calculation_tracker.track_ratio_calculation(
                                    "D/E Ratio",
                                    result.value,
                                    {'BS': bs_context},
                                    result.inputs_used
                                )                                

                    except Exception as e:
                        log_debug(f"Error calculating D/E Ratio: {str(e)}", include_trace=True)      
                        
                    try:
                        # Equity Multiplier
                        equity_multiplier_results = self.calculate_equity_multiplier(bs_context)
                        if equity_multiplier_results:
                            for result in equity_multiplier_results:
                                result_set.add_result(result)
                                log_debug(f"Added Equity Multiplier: {result.value:.4f}")
                                self.calculation_tracker.track_ratio_calculation(
                                    "Equity Multiplier",
                                    result.value,
                                    {'BS': bs_context},
                                    result.inputs_used
                                )                                

                    except Exception as e:
                        log_debug(f"Error calculating Equity Multiplier: {str(e)}", include_trace=True)                                                         
                        
                    try:
                        # Quick Ratio
                        quick_ratio_results = self.calculate_quick_ratio(bs_context)
                        if quick_ratio_results:
                            for result in quick_ratio_results:
                                result_set.add_result(result)
                                log_debug(f"Added Quick Ratio: {result.value:.4f}")
                                self.calculation_tracker.track_ratio_calculation(
                                    "Quick Ratio",
                                    result.value,
                                    {'BS': bs_context},
                                    result.inputs_used
                                )                                

                    except Exception as e:
                        log_debug(f"Error calculating Quick Ratio: {str(e)}", include_trace=True)      
                        
                        
                    try:
                        # Debt to Capital Ratio
                        debt_to_capital_ratio_results = self.calculate_debt_to_capital_ratio(bs_context)
                        if debt_to_capital_ratio_results:
                            for result in debt_to_capital_ratio_results:
                                result_set.add_result(result)
                                log_debug(f"Added Debt to Capital Ratio: {result.value:.4f}")
                                self.calculation_tracker.track_ratio_calculation(
                                    "Debt to Capital Ratio",
                                    result.value,
                                    {'BS': bs_context},
                                    result.inputs_used
                                )                                

                    except Exception as e:
                        log_debug(f"Error calculating Debt to Capital Ratio: {str(e)}", include_trace=True)                            
                    
                        # Equity Ratio
                    try:
                        equity_ratio_results = self.calculate_equity_ratio(bs_context)
                        if equity_ratio_results:
                            for result in equity_ratio_results:
                                result_set.add_result(result)
                                log_debug(f"Added Debt to Equity Ratio: {result.value:.4f}")
                                self.calculation_tracker.track_ratio_calculation(
                                    "Equity Ratio",
                                    result.value,
                                    {'BS': bs_context},
                                    result.inputs_used
                                )                                

                    except Exception as e:
                        log_debug(f"Error calculating Equity Ratio: {str(e)}", include_trace=True)

                        # Dept to Assets Ratio
                    try:
                        debt_to_assets_results = self.calculate_debt_to_assets(bs_context)
                        if debt_to_assets_results:
                            for result in debt_to_assets_results:
                                result_set.add_result(result)
                                log_debug(f"Added Debt to Dept to Assets Ratio: {result.value:.4f}")
                                self.calculation_tracker.track_ratio_calculation(
                                    "Dept to Assets Ratio",
                                    result.value,
                                    {'BS': bs_context},
                                    result.inputs_used
                                )                   
                    except Exception as e:
                        log_debug(f"Error calculating Debt to Assets: {str(e)}", include_trace=True)

                        # Cash Ratio
                    try:
                        cash_ratio_results = self.calculate_cash_ratio(bs_context)
                        if cash_ratio_results:
                            for result in cash_ratio_results:
                                result_set.add_result(result)
                                log_debug(f"Added Debt to Cash Ratio: {result.value:.4f}")
                                self.calculation_tracker.track_ratio_calculation(
                                    "Cash Ratio",
                                    result.value,
                                    {'BS': bs_context},
                                    result.inputs_used
                                )                   
                    except Exception as e:
                        log_debug(f"Error calculating Cash Ratio: {str(e)}", include_trace=True)                    
                    
                    # try:
                    #     # Operating Cash Flow Ratio (needs CF context)
                    #     cf_context = result_set.get_matching_context(bs_context, 'CF')
                    #     if cf_context:
                    #         ocf_ratio_results = self.calculate_operating_cash_flow_ratio(bs_context, cf_context)
                    #         if ocf_ratio_results:
                    #             for result in ocf_ratio_results:
                    #                 result_set.add_result(result)
                    #                 log_debug(f"Added Operating CF Ratio: {result.value:.4f}")
                    # except Exception as e:
                    #     log_debug(f"Error calculating Operating CF Ratio: {str(e)}", include_trace=True)          
                        
                        
                    try:
                        # Net Debt to Equity
                        net_debt_equity_results = self.calculate_net_debt_to_equity(bs_context)
                        if net_debt_equity_results:
                            for result in net_debt_equity_results:
                                result_set.add_result(result)
                                log_debug(f"Added Net Debt to Equity: {result.value:.4f}")
                                
                    except Exception as e:
                        log_debug(f"Error calculating Net Debt to Equity: {str(e)}", include_trace=True)

                    try:
                        # Long-term Debt Ratio
                        ltd_ratio_results = self.calculate_long_term_debt_ratio(bs_context)
                        if ltd_ratio_results:
                            for result in ltd_ratio_results:
                                result_set.add_result(result)
                                log_debug(f"Added Long-term Debt Ratio: {result.value:.4f}")
                                
                    except Exception as e:
                        log_debug(f"Error calculating Long-term Debt Ratio: {str(e)}", include_trace=True)

                    try:
                        # Net Debt to Capital
                        net_debt_capital_results = self.calculate_net_debt_to_capital(bs_context)
                        if net_debt_capital_results:
                            for result in net_debt_capital_results:
                                result_set.add_result(result)
                                log_debug(f"Added Net Debt to Capital: {result.value:.4f}")
                                
                    except Exception as e:
                        log_debug(f"Error calculating Net Debt to Capital: {str(e)}", include_trace=True)                                  
                                                    
                #################################################  CF Ratios
                if cf_context:  # Only calculate CF ratios if context available
                    log_debug(f"Entering cf_context calculation in calculate_ratios_for_contexts with: {cf_context}")
                    
                    try:
                        # Operating Cash Flow Ratio
                        ocf_ratio_results = self.calculate_operating_cash_flow_ratio(bs_context, cf_context)
                        if ocf_ratio_results:
                            for result in ocf_ratio_results:
                                result_set.add_result(result)
                                log_debug(f"Added Operating CF Ratio: {result.value:.4f}")
                                self.calculation_tracker.track_ratio_calculation(
                                    "Operating CF Ratio",
                                    result.value,
                                    {'BS': bs_context, 'CF': cf_context},
                                    result.inputs_used   
                                )                             
                    except Exception as e:
                        log_debug(f"Error calculating Operating CF Ratio: {str(e)}", include_trace=True)
                    
                    try:
                        # Operating Cash Flow Margin
                        if is_context:  # Need both CF and IS contexts
                            ocf_margin_results = self.calculate_operating_cash_flow_margin(cf_context, is_context)
                            if ocf_margin_results:
                                for result in ocf_margin_results:
                                    result_set.add_result(result)
                                    log_debug(f"Added Operating CF Margin: {result.value:.4f}")
                    except Exception as e:
                        log_debug(f"Error calculating Operating CF Margin: {str(e)}", include_trace=True)
                   
                                   
                #########################################################################################
                # Calculate IS ratios
                if is_context:
                    log_debug("\nCalculating IS ratios...")
                    
                    try:
                        # Net Margin
                        net_margin_results = self.calculate_net_margin(is_context)
                        if net_margin_results:
                            for result in net_margin_results:
                                result_set.add_result(result)
                                log_debug(f"Added Net Margin: {result.value:.4f}")
                                
                                # Track ratio calculation
                                self.calculation_tracker.track_ratio_calculation(
                                    "Net Margin",
                                    result.value,
                                    {'IS': is_context},
                                    result.inputs_used
                                )                                      
                        
                    except Exception as e:
                        log_debug(f"Error calculating Net Margin: {str(e)}", include_trace=True)                           
                    
                    try:
                        # Gross Margin
                        gross_margin_results = self.calculate_gross_margin(is_context)
                        if gross_margin_results:
                            for result in gross_margin_results:
                                result_set.add_result(result)
                                log_debug(f"Added Gross Margin: {result.value:.4f}")
                                self.calculation_tracker.track_ratio_calculation(
                                    "Gross Margin",
                                    result.value,
                                    {'IS': is_context},
                                    result.inputs_used
                                )       
                                            
                    except Exception as e:
                        log_debug(f"Error calculating Gross Margin: {str(e)}", include_trace=True)       

                    try:
                        # Cost of Sales
                        cost_of_sales_results = self.calculate_cost_of_sales_ratio(is_context)
                        if cost_of_sales_results:
                            for result in cost_of_sales_results:
                                result_set.add_result(result)
                                log_debug(f"Added Cost of Sales: {result.value:.4f}")
                                self.calculation_tracker.track_ratio_calculation(
                                    "Cost of Sales",
                                    result.value,
                                    {'IS': is_context},
                                    result.inputs_used
                                )       
                                            
                    except Exception as e:
                        log_debug(f"Error calculating Cost of Sales: {str(e)}", include_trace=True)       
                        
                        
                    try:
                        # EBIT Margin
                        ebit_margin_results = self.calculate_ebit_margin(is_context)
                        if ebit_margin_results:
                            for result in ebit_margin_results:
                                result_set.add_result(result)
                                log_debug(f"Added EBIT Margin: {result.value:.4f}")
                    except Exception as e:
                        log_debug(f"Error calculating EBIT Margin: {str(e)}", include_trace=True)

                    try:
                        # EBITDA Margin
                        ebitda_margin_results = self.calculate_ebitda_margin(is_context)
                        if ebitda_margin_results:
                            for result in ebitda_margin_results:
                                result_set.add_result(result)
                                log_debug(f"Added EBITDA Margin: {result.value:.4f}")
                    except Exception as e:
                        log_debug(f"Error calculating EBITDA Margin: {str(e)}", include_trace=True)

                    try:
                        # Interest Coverage
                        coverage_results = self.calculate_interest_coverage(is_context)
                        if coverage_results:
                            for result in coverage_results:
                                result_set.add_result(result)
                                log_debug(f"Added Interest Coverage: {result.value:.4f}")
                    except Exception as e:
                        log_debug(f"Error calculating Interest Coverage: {str(e)}", include_trace=True)                        

                    # Operating Profit Margin = EBIT Margin
                    # try:
                    #     opm_results = self.calculate_operating_profit_margin(is_context)
                    #     if opm_results:
                    #         for result in opm_results:
                    #             result_set.add_result(result)
                    #             log_debug(f"Added Operating Profit Margin: {result.value:.4f}")
                    # except Exception as e:
                    #     log_debug(f"Error calculating Operating Profit Margin: {str(e)}", include_trace=True)
                    
                    try:
                        # Dividend Payout
                        dividend_payout_results = self.calculate_dividend_payout(is_context)
                        if dividend_payout_results:
                            for result in dividend_payout_results:
                                result_set.add_result(result)
                                log_debug(f"Added Dividend Payout: {result.value:.4f}")
                    except Exception as e:
                        log_debug(f"Error calculating Dividend Payout: {str(e)}", include_trace=True)                    

                #################################################                                        
                # Calculate mixed ratios when both contexts available
                if bs_context and is_context:
                    log_debug("\nCalculating mixed ratios...")
                    
                    try:
                        # Return on Assets
                        roa_results = self.calculate_return_on_assets(bs_context, is_context)
                        if roa_results:
                            for result in roa_results:
                                result_set.add_result(result)
                                log_debug(f"Added ROA: {result.value:.4f}")
                    except Exception as e:
                        log_debug(f"Error calculating ROA: {str(e)}", include_trace=True)                                       
                
                    # ROE
                    try:
                        roe_results = self.calculate_return_on_equity(bs_context, is_context)
                        if roe_results:
                            for result in roe_results:
                                result_set.add_result(result)
                                log_debug(f"Added ROE: {result.value:.4f}")
                    except Exception as e:
                        log_debug(f"Error calculating ROE: {str(e)}", include_trace=True)                
                
                    # ROCE
                    try:
                        roce_results = self.calculate_roce(bs_context, is_context)
                        if roce_results:
                            for result in roce_results:
                                result_set.add_result(result)
                                log_debug(f"Added ROCE: {result.value:.4f}")
                    except Exception as e:
                        log_debug(f"Error calculating ROCE: {str(e)}", include_trace=True)

                    # Asset Turnover
                    try:
                        at_results = self.calculate_asset_turnover(bs_context, is_context)
                        if at_results:
                            for result in at_results:
                                result_set.add_result(result)
                                log_debug(f"Added Asset Turnover: {result.value:.4f}")
                    except Exception as e:
                        log_debug(f"Error calculating Asset Turnover: {str(e)}", include_trace=True)                
                
                    # try:
                    #     # Operating Cash Flow Margin
                    #     if cf_context and is_context:
                    #         ocf_margin_results = self.calculate_operating_cash_flow_margin(cf_context, is_context)
                    #         if ocf_margin_results:
                    #             for result in ocf_margin_results:
                    #                 result_set.add_result(result)
                    #                 log_debug(f"Added Operating CF Margin: {result.value:.4f}")
                    # except Exception as e:
                    #     log_debug(f"Error calculating Operating CF Margin: {str(e)}", include_trace=True)

                    try:
                        # Inventory Turnover
                        inventory_turnover_results = self.calculate_inventory_turnover(bs_context, is_context)
                        if inventory_turnover_results:
                            for result in inventory_turnover_results:
                                result_set.add_result(result)
                                log_debug(f"Added Inventory Turnover: {result.value:.4f}")
                    except Exception as e:
                        log_debug(f"Error calculating Inventory Turnover: {str(e)}", include_trace=True)
                        
                    try:
                        # Net Debt to EBITDA
                        net_debt_ebitda_results = self.calculate_net_debt_to_ebitda(bs_context, is_context)
                        if net_debt_ebitda_results:
                            for result in net_debt_ebitda_results:
                                result_set.add_result(result)
                                log_debug(f"Added Net Debt to EBITDA: {result.value:.4f}")
                                
                    except Exception as e:
                        log_debug(f"Error calculating Net Debt to EBITDA: {str(e)}", include_trace=True)

                    try:
                        # Total Debt to EBITDA
                        total_debt_ebitda_results = self.calculate_total_debt_to_ebitda(bs_context, is_context)
                        if total_debt_ebitda_results:
                            for result in total_debt_ebitda_results:
                                result_set.add_result(result)
                                log_debug(f"Added Total Debt to EBITDA: {result.value:.4f}")
                                
                    except Exception as e:
                        log_debug(f"Error calculating Total Debt to EBITDA: {str(e)}", include_trace=True)
                        
                                        
                # ###############################
                # Add to ratio analysis regardless of any calculation failures
                try:
                    self.ratio_analysis.add_result_set(result_set)
                    log_debug(f"\nAdded {len(result_set.results)} ratios to analysis")
                except Exception as e:
                    log_debug(f"Error adding results to analysis: {str(e)}", include_trace=True)
                
                # Complete calculation tracking
                self.calculation_tracker.complete_calculation(None, "completed")
                self.calculation_tracker.summarize_calculations()
                
                return result_set
                
            except Exception as e:
                log_debug(f"Error calculating ratios: {str(e)}", include_trace=True)
                self.calculation_tracker.complete_calculation(None, "error", str(e))
                return result_set
                
        except Exception as e:
            log_debug(f"Critical error in ratio calculation: {str(e)}", include_trace=True)
            return RatioResultSet(None, self.data_manager)    



class DebtCalculator:
    """Enhanced calculator for determining total debt by considering all sources"""
    
    def __init__(self, data_manager):
        self.data_manager = data_manager
        self.logger = logger
        
    def calculate_total_debt(self, period: tuple) -> tuple:
        """Calculate total debt using data manager and component tracking"""
        log_debug("\n=== Calculating Total Debt ===")
        
        try:
            # Try direct total debt value through data manager
            value, metadata = self.data_manager.get_value('total_debt', period, 'BS')
            if value is not None:
                log_debug(f"Found direct total debt value: {value}")
                return value, metadata
            
            # Calculate from components using concept maps
            total = 0
            components = {}
            
            # Get components from concept map
            concept_info = self.data_manager.fact_provider.concept_maps.get('total_debt')
            if not concept_info:
                log_debug("No concept mapping found for total_debt")
                return None, None

            # Get components from ComponentRegistry
            debt_components = self.data_manager.component_registry.get_components('total_debt')
            if not debt_components:
                log_debug("No components defined in registry for total_debt")
                return None, None
            
            log_debug(f"Calculating from components: {debt_components}")
            
            # Get each component through data manager
            for component in debt_components:
                value, metadata = self.data_manager.get_value(component, period, 'BS')
                if value is not None:
                    log_debug(f"Found component {component}: {value}")
                    total += value
                    components[component] = {
                        'value': value,
                        'metadata': metadata
                    }
                    
                    # Track inputs through input tracker
                    if hasattr(self.data_manager.fact_provider, 'ratio_calculator'):
                        self.data_manager.fact_provider.ratio_calculator.input_tracker.add_input(
                            name=component,
                            qname=metadata.qname if metadata else 'Unknown',
                            value=value,
                            context_id=metadata.context_id if metadata else None,
                            statement='BS',
                            indent_level=1
                        )
            
            if components:
                # Create metadata for compound value
                metadata = ValueMetadata(
                    qname="Calculated",
                    source_statement='BS',
                    is_custom=all(c['metadata'].is_custom for c in components.values()),
                    components=components
                )
                
                log_debug(f"Calculated total debt: {total}")
                return total, metadata
                
            return None, None
                
        except Exception as e:
            log_debug(f"Error calculating total debt: {str(e)}", include_trace=True)
            return None, None

 

class DebtEquityCalculator:
    def __init__(self, fact_provider):
        self.fact_provider = fact_provider
        self.data_manager = fact_provider.financial_data_manager
        self.logger = logger
        
    def calculate_total_debt(self, period: tuple) -> tuple:
        """Calculate total debt using data manager and tracking components"""
        log_debug("\n=== Calculating Total Debt with Component Tracking ===")
        
        # Try direct value through data manager
        value, metadata = self.data_manager.get_value('total_debt', period, 'BS')
        if value is not None:
            log_debug("Found direct total debt value")
            return value, metadata
            
        # Calculate from components
        components = {}
        total = 0
        
        # Define debt components
        debt_components = [
            'long_term_borrowings',
            'short_term_borrowings',
            'lease_liabilities',
            'bonds_payable',
            'bank_loans'
        ]
        
        # Get each component through data manager
        for component in debt_components:
            value, metadata = self.data_manager.get_value(component, period, 'BS')
            if value is not None:
                log_debug(f"Found component {component}: {value}")
                total += value
                components[component] = {
                    'value': value,
                    'metadata': metadata
                }
        
        if components:
            # Create metadata for the compound value
            metadata = ValueMetadata(
                qname="Calculated",
                source_statement='BS',
                is_custom=all(c['metadata'].is_custom for c in components.values()),
                components=components
            )
            
            log_debug(f"Calculated total debt: {total} from {len(components)} components")
            return total, metadata
            
        return None, None
        
    def calculate_total_equity(self, period: tuple) -> tuple:
        """Calculate total equity using data manager"""
        log_debug("\n=== Calculating Total Equity ===")
        
        # Try direct equity value
        value, metadata = self.data_manager.get_value('total_equity', period, 'BS')
        if value is not None:
            log_debug(f"Found direct total equity value: {value}")
            return value, metadata
        
        # If direct value not found, try components
        components = {}
        total = 0
        
        # Define equity components
        equity_components = [
            'share_capital',
            'retained_earnings',
            'reserves'
        ]
        
        # Get each component through data manager
        for component in equity_components:
            value, metadata = self.data_manager.get_value(component, period, 'BS')
            if value is not None:
                log_debug(f"Found component {component}: {value}")
                total += value
                components[component] = {
                    'value': value,
                    'metadata': metadata
                }
        
        if components:
            # Create metadata for the compound value
            metadata = ValueMetadata(
                qname="Calculated",
                source_statement='BS',
                is_custom=all(c['metadata'].is_custom for c in components.values()),
                components=components
            )
            
            log_debug(f"Calculated total equity: {total} from {len(components)} components")
            return total, metadata
            
        return None, None
   




class TotalCapitalCalculator:
    """Dedicated calculator for determining total capital with multiple approaches"""
    
    def __init__(self, fact_provider):
        self.fact_provider = fact_provider
        self.logger = logger
        
    def calculate_total_capital(self, period: tuple, dependency_chain=None) -> tuple:
        """Calculate total capital using all available methods"""
        log_debug("\n=== Calculating Total Capital ===")
        
        # Try direct value first using the extensive concept mapping
        direct_result = self._try_direct_value(period)
        if direct_result[0] is not None:
            log_debug("Found direct total capital value")
            return direct_result
            
        # If direct lookup fails, try calculation methods in sequence
        # Pass dependency_chain to prevent circular references
        total_debt = self.fact_provider.get_compound_value('total_debt', period, 'BS', dependency_chain)
        if total_debt[0] is None:
            log_debug("Could not calculate - total_debt not available")
            return None, None, False
            
        total_equity = self.fact_provider.get_fact_value('total_equity', period, 'BS')
        if total_equity[0] is None:
            log_debug("Could not calculate - total_equity not available")
            return None, None, False
        
        value = total_debt[0] + total_equity[0]
        log_debug(f"Calculated total capital = {total_debt[0]} (debt) + {total_equity[0]} (equity) = {value}")
        
        # Use non-custom qname if available
        qname = total_equity[1] if not total_equity[2] else total_debt[1]
        is_custom = total_equity[2] and total_debt[2]
        
        return value, qname, is_custom
        
    def _try_direct_value(self, period: tuple) -> tuple:
        """Try to get direct total capital value"""
        log_debug("\nTrying direct total capital lookup")
        
        # Try primary concept using existing get_fact_value
        result = self.fact_provider.get_fact_value('total_capital', period, 'BS')
        if result[0] is not None:
            log_debug("Found direct total capital value")
            return result
            
        # Try alternative concepts
        alternatives = [
            'capital_and_reserves',
            'equity_and_borrowings',
            'total_capitalisation'
        ]
        
        for concept in alternatives:
            result = self.fact_provider.get_fact_value(concept, period, 'BS')
            if result[0] is not None:
                log_debug(f"Found total capital via alternative concept: {concept}")
                return result
                
        log_debug("No direct total capital value found")
        return None, None, False

    
 
      
class FilingAnalyzer:
    """Complete system for analyzing financial statement filings"""
    
    # Statement type constants
    COVER    = "1Cover"
    STMTS    = "2Financial Statements"  
    NOTES    = "3Notes to Financial Statements"
    POLICIES = "4Accounting Policies"
    TABLES   = "5Notes Tables"
    DETAILS  = "6Notes Details"
    UNCATEG  = "7Uncategorized"
    
    # Role type patterns
    ROLE_PATTERNS = {
        'STMT': r".* - statement - ",
        'DOC': r".* - document - ",
        'DISCLOSURE': r".* - disclosure - ",
        'DETAILS': r"(?!.*details)",
        'COMPREHENSIVE': r"(?!.*comprehensive)",
        'IS_COMPREHENSIVE': r"(?=.*comprehensive)",
        'PARENTHETICAL': r"pa?r[ae]ne?th\w?[aei]+\w?t?h?i?c",
    }


    def __init__(self, model_xbrl):
        self.model_xbrl = model_xbrl
        self.statement_detector = StatementTypeDetector()
        log_debug("\n=== Initializing FilingAnalyzer ===")
        self.filing_type = self._detect_filing_type()
        # Initialize contexts and mappings for proper Arelle GUI integration
        self.statement_periods = self._initialize_statement_contexts()
        self.context_mappings = self._create_context_mappings(self.statement_periods)
        log_debug(f"Initialized with {len(self.statement_periods)} statement types")
        
        # # Initialize mappings without Arelle GUI integration
        # self.context_mappings = {}            
        # # just empty init
        # self.statement_periods = {
        #     'BS': {},
        #     'IS': {},
        #     'CI': {},
        #     'CF': {},
        #     'EQ': {}
        # }        
                        
    def _detect_filing_type(self) -> str:
        """Detect filing type based on taxonomy and patterns"""
        log_debug("\nAnalyzing namespaces for filing type detection:")
        
        # IFRS/ESEF namespace patterns
        ifrs_patterns = [
            r"http://xbrl.ifrs.org/.*",
            r"http://www.esma.europa.eu/.*",
            r".*/ifrs/.*",
            r".*/fr/.*",
            r".*/gaap/ifrs/.*"
        ]
        
        # Log namespace analysis
        for ns in self.model_xbrl.namespaceDocs.keys():
            if ns:
                log_debug(f"Found namespace: {ns}")
                if any(re.match(pattern, ns) for pattern in ifrs_patterns):
                    log_debug("Detected IFRS filing based on namespace")
                    return "IFRS"
                    
        # Check for JP-FSA
        if "jp-fsa" in self.model_xbrl.modelManager.disclosureSystem.names:
            log_debug("Detected JP-FSA filing")
            return "JP-FSA"
            
        # Check for US-GAAP patterns in role definitions
        us_gaap_pattern = re.compile(r"([0-9]+) - (Statement|Disclosure|Schedule|Document) - (.+)")
        role_definitions = [
            self.model_xbrl.roleTypeDefinition(roleURI)
            for roleURI in self.model_xbrl.relationshipSet(XbrlConst.parentChild).linkRoleUris
        ]
        
        if any(us_gaap_pattern.match(rd) for rd in role_definitions if rd):
            log_debug("Detected US-GAAP filing based on role patterns")
            return "US-GAAP"
            
        log_debug("No specific filing type detected, defaulting to IFRS")
        return "IFRS"

    def _initialize_statement_contexts(self) -> dict:
        """Initialize statement contexts with mapping"""
        statement_contexts = {
            'BS': {},
            'IS': {},
            'CI': {},
            'CF': {},
            'EQ': {}
        }
        
        log_debug("\n=== Initializing Statement Contexts (FilingAnalyzer._initialize_statement_contexts) ===")
        
        try:
            relationship_set = self.model_xbrl.relationshipSet(XbrlConst.parentChild)            
            if not relationship_set:
                log_debug("No relationship set found")
                return statement_contexts

            # # Debug role types
            # log_debug("\nChecking role types:")
            # for roleURI in relationship_set.linkRoleUris:
            #     role_types = self.model_xbrl.roleTypes.get(roleURI, [])
            #     if role_types:
            #         role_type = role_types[0]
            #         stmt_type = getattr(role_type, '_statementType', '')
            #         log_debug(f"\nRole URI: {roleURI}")
            #         log_debug(f"Statement Type: {stmt_type}")
            #         facts = getattr(role_type, '_tableFacts', set())
            #         log_debug(f"Number of facts: {len(facts)}")

            context_facts = defaultdict(list)
            for fact in self.model_xbrl.facts:
                if fact.context is not None:
                    context_facts[fact.context.id].append(fact)
                        
            # Process roles to determine statement types and contexts
            log_debug("\nProcessing roles for contexts:")
            for roleURI in relationship_set.linkRoleUris:
                role_types = self.model_xbrl.roleTypes.get(roleURI, [])
                if not role_types:
                    continue
                        
                role_type = role_types[0]
                stmt_type = getattr(role_type, '_statementType', '').rstrip('P')
                
                log_debug(f"\nProcessing role type:")
                log_debug(f"Statement Type: {stmt_type}")
                
                if stmt_type not in statement_contexts:
                    log_debug(f"Statement type {stmt_type} not in tracked types")
                    continue
                        
                facts = getattr(role_type, '_tableFacts', set())
                log_debug(f"Found {len(facts)} facts")
                
                for fact in facts:
                    if fact.context is None:
                        continue
                            
                    context = fact.context
                    context_id = context.id
                        
                    # Store context with its period info
                    if context.isInstantPeriod:
                        period = (None, context.instantDatetime)
                    else:
                        period = (context.startDatetime, context.endDatetime)
                            
                    if period[1]:  # Only store if we have an end date
                        statement_contexts[stmt_type][context_id] = period
                        # log_debug(f"Added context {context_id} to {stmt_type}")
                            
            # Log context summary
            log_debug("\nFinal context counts:")
            for stmt_type, contexts in statement_contexts.items():
                log_debug(f"{stmt_type}: {len(contexts)} contexts")
                if contexts:
                    log_debug("Sample contexts:")
                    for ctx, period in list(contexts.items())[:3]:
                        log_debug(f"  Context {ctx}: "
                                f"Start={period[0].strftime('%Y-%m-%d') if period[0] else 'None'}, "
                                f"End={period[1].strftime('%Y-%m-%d')}")
                            
            return statement_contexts
                    
        except Exception as e:
            log_debug(f"Error initializing contexts: {str(e)}", include_trace=True)
            return statement_contexts

    def _create_context_mappings(self, statement_periods: dict) -> dict:
        """Create mappings between contexts with matching end dates"""
        mappings = {}
        log_debug("\n=== Creating Context Mappings (FilingAnalyzer._create_context_mappings)===")
        
        # Start with BS contexts as anchors
        for bs_context, bs_period in statement_periods['BS'].items():
            if not bs_period[1]:  # Skip if no end date
                continue
                
            log_debug(f"\nProcessing BS context: {bs_context}")
            log_debug(f"Period: {bs_period}")
            
            mappings[bs_context] = {
                'period': bs_period,
                'matches': {'IS': [], 'CI': [], 'CF': [], 'EQ': []}
            }
            
            # Find matching contexts in other statements
            for stmt_type in ['IS', 'CI', 'CF', 'EQ']:
                for other_context, other_period in statement_periods[stmt_type].items():
                    if other_period[1] == bs_period[1]:
                        mappings[bs_context]['matches'][stmt_type].append(other_context)
                        log_debug(f"Matched with {stmt_type} context: {other_context}")
        
        # Log summary
        log_debug("\nMapping Summary:")
        for bs_ctx, mapping in mappings.items():
            log_debug(f"\nBS Context: {bs_ctx}")
            log_debug(f"Period: {mapping['period']}")
            for stmt_type, contexts in mapping['matches'].items():
                log_debug(f"{stmt_type} matches: {contexts}")
        
        return mappings
     
    def get_facts_for_role(self, role_type) -> set:
        """Get facts for a specific role with proper hierarchy information"""
        facts = set()
        
        try:
            relSet = self.model_xbrl.relationshipSet(XbrlConst.parentChild, role_type.roleURI)            
            if not relSet:
                return facts

            # Track facts in hierarchical order
            def process_relationships(concepts, parent=None, level=0):
                for concept in concepts:
                    if isinstance(concept, ModelConcept):
                        for fact in self.model_xbrl.factsByQname.get(concept.qname, []):
                            if fact.context is not None:
                                fact._tableLevel = level  # Set hierarchy level
                                fact._tableParent = parent  # Set parent relationship
                                facts.add(fact)
                        # Process children using relationship set
                        child_rels = relSet.fromModelObject(concept)
                        if child_rels:
                            child_concepts = [rel.toModelObject for rel in child_rels]
                            process_relationships(child_concepts, concept, level + 1)

            # Get root relationships and extract their concepts
            root_concepts = relSet.rootConcepts
            process_relationships(root_concepts)

            log_debug(f"\n=== Processing Facts for Role ===")
            log_debug(f"Role Definition: {role_type.definition}")
            log_debug(f"Statement Type: {getattr(role_type, '_statementType', 'None')}")

            # Count facts per context for filtering
            context_fact_counts = defaultdict(int)
            for fact in facts:
                context_fact_counts[fact.context.id] += 1

            # Keep facts from main contexts
            filtered_facts = set()
            for context_id, count in context_fact_counts.items():
                if count >= 5:  # Only keep contexts with significant facts
                    log_debug(f"Including context {context_id} with {count} facts")
                    filtered_facts.update(f for f in facts if f.context.id == context_id)

            log_debug(f"Found {len(filtered_facts)} facts in {len([c for c,n in context_fact_counts.items() if n >= 5])} main contexts")
            return filtered_facts

        except Exception as e:
            log_debug(f"Error processing facts: {str(e)}", include_trace=True)
            return facts

    def sort_statements(self, statements: List[Tuple[str, object]]) -> List[Tuple[str, object]]:
        """Sort statements based on type and filing standard"""
        def get_sort_key(stmt):
            role_type = stmt[1]
            stmt_type = getattr(role_type, '_statementType', None)
            
            # Define order of statements
            type_order = {
                'BS': 0,   # Balance Sheet
                'BSP': 0,  # Balance Sheet (parenthetical)
                'IS': 1,   # Income Statement
                'ISP': 1,  # Income Statement (parenthetical)
                'CI': 2,   # Comprehensive Income
                'CIP': 2,  # Comprehensive Income (parenthetical)
                'EQ': 3,   # Changes in Equity
                'EQP': 3,  # Equity (parenthetical)
                'CF': 4,   # Cash Flow
                'CFP': 4,  # Cash Flow (parenthetical)
                'OTH': 5   # Other financial statements
            }
            
            order = type_order.get(stmt_type, 999) if stmt_type else 999
            # For statements of same type, non-parenthetical before parenthetical
            is_parenthetical = stmt_type and stmt_type.endswith('P') if stmt_type else False
            
            return (order, is_parenthetical, stmt[0])
        
        return sorted(
            [s for s in statements if s[1]._tableIndex[0] == self.STMTS],
            key=get_sort_key
        )

    def analyze_role_type(self, definition: str, role_type) -> None:
        """Two-phase analysis for proper Arelle table detection"""
        log_debug(f"\n=== Analyzing Role Type (Two-Phase) ===")
        log_debug(f"Definition: {definition}")

        # Phase 1: Initial table code evaluation
        role_info = self.get_role_info(definition)
        initial_group = self.determine_initial_group(role_info)
        log_debug(f"Initial group: {initial_group}")

        # Get statement type with confidence
        element_count = len(self.model_xbrl.relationshipSet(
            XbrlConst.parentChild, role_type.roleURI).modelRelationships)
        stmt_type, confidence = self.statement_detector.detect_statement_type(
            definition, element_count)
        log_debug(f"Detected statement type: {stmt_type} (confidence: {confidence})")

        # Phase 2: Re-evaluate group based on statement detection
        if stmt_type and confidence > 0.4:
            if stmt_type in ['BS', 'BSP', 'IS', 'ISP', 'CI', 'CIP', 'EQ', 'EQP', 'CF', 'CFP']:
                log_debug("Updating group to Financial Statements based on detection")
                final_group = self.STMTS
            else:
                final_group = initial_group
        else:
            final_group = initial_group

        log_debug(f"Final group: {final_group}")

        # Set sequence number
        sequence = str(get_statement_type_order().get(stmt_type, 99)).zfill(2) if stmt_type else "99"

        # Set essential properties for Arelle GUI
        role_type._initialGroup = initial_group
        role_type._statementType = stmt_type
        role_type._tableCode = stmt_type  # Critical for table detection
        role_type._confidence = confidence
        role_type._tableIndex = (final_group, sequence, definition)
        role_type._tableFacts = self.get_facts_for_role(role_type)

        # Set up contexts and periods for GUI
        if stmt_type:
            contexts = {}
            for fact in role_type._tableFacts:
                if fact.context is not None:
                    context_id = fact.context.id
                    if fact.context.isInstantPeriod:
                        period = (None, fact.context.instantDatetime)
                    else:
                        period = (fact.context.startDatetime, fact.context.endDatetime)
                    if period[1]:  # Only store if we have an end date
                        contexts[context_id] = period

            # Set view properties required by Arelle GUI
            role_type._tableContexts = contexts
            role_type._viewProperties = {
                'label': definition,
                'roleURI': role_type.roleURI,
                'tableCode': stmt_type,
                'tableIndex': sequence,
                'periods': list(contexts.values())
            }

            log_debug(f"View properties set:")
            log_debug(f"- Table Code: {stmt_type}")
            log_debug(f"- Statement Type: {stmt_type}")
            log_debug(f"- Contexts: {len(contexts)}")
    
    def determine_initial_group(self, role_info: dict) -> str:
        """First-pass grouping based on English patterns"""
        if role_info['is_document']:
            return self.COVER
        if role_info['is_statement']:
            return self.STMTS  # Initial categorization as financial statement
        if role_info['is_disclosure']:
            if role_info['has_details']:
                return self.DETAILS
            return self.NOTES
        return self.UNCATEG

    def determine_final_group(self, stmt_type: str, initial_group: str, confidence: float) -> str:
        """
        Final grouping that allows detailed detection to override initial group
        when we're confident about the statement type
        """
        log_debug(f"\nDetermining final group:")
        log_debug(f"Statement type received: {stmt_type}")
        log_debug(f"Initial group: {initial_group}")
        log_debug(f"Confidence: {confidence}")

        # If we found a specific statement type with good confidence,
        # it should override the initial group
        if stmt_type in ['BS', 'BSP', 'IS', 'ISP', 'CI', 'CIP', 'EQ', 'EQP', 'CF', 'CFP'] and confidence > 0.4:
            log_debug("Valid statement type found - setting to Financial Statements")
            return self.STMTS
                
        # Otherwise keep the initial categorization
        log_debug(f"Keeping initial group: {initial_group}")
        return initial_group

    def initialize_contexts(self):
        """Initialize contexts after table facts have been assigned"""
        log_debug("\n=== Initializing Statement Contexts (FilingAnalyzer.initialize_contexts) ===")
        self.statement_periods = self._initialize_statement_contexts()
        populated_statements = {stmt: len(contexts) 
                            for stmt, contexts in self.statement_periods.items() 
                            if contexts}
        log_debug("\nPopulated statement contexts:")
        for stmt, count in populated_statements.items():
            log_debug(f"- {stmt}: {count} contexts")
   
    def update_contexts_from_merged_statements(self, merged_statements):
        """Update contexts after statement merging"""
        log_debug("\n=== Updating Contexts from Merged Statements ===")
        
        # Reset statement periods
        self.statement_periods = {
            'BS': {},
            'IS': {},
            'CI': {},
            'CF': {},
            'EQ': {}
        }
        
        # Process each merged statement
        for stmt in merged_statements:
            stmt_type = stmt.get('statement_type')
            if stmt_type in self.statement_periods:
                # Get facts from merged statement
                facts = getattr(stmt.get('role_type'), '_tableFacts', set())
                if facts:
                    log_debug(f"\nProcessing {stmt_type} statement facts")
                    # Add contexts from facts
                    for fact in facts:
                        if fact.context:
                            context_id = fact.context.id
                            if fact.context.isInstantPeriod:
                                period = (None, fact.context.instantDatetime)
                            else:
                                period = (fact.context.startDatetime, fact.context.endDatetime)
                            if period[1]:  # Only store if we have an end date
                                self.statement_periods[stmt_type][context_id] = period
                                
        # Log updated context counts
        for stmt_type, contexts in self.statement_periods.items():
            log_debug(f"{stmt_type}: {len(contexts)} contexts")

    def prepare_role_type_view(self, role_type):
        """Prepare role type for Arelle view system"""
        if hasattr(role_type, '_tableIndex'):
            # Set view properties directly instead of trying to set propertyView
            role_type._viewProperties = {
                "label": role_type.definition,
                "roleURI": role_type.roleURI,
                "tableCode": getattr(role_type, '_tableCode', ''),
                "tableIndex": role_type._tableIndex[1] if role_type._tableIndex else ''
            }
            
            # Set table codes for view system
            if hasattr(role_type, '_statementType'):
                role_type._tableCodes = [role_type._statementType]

    def _initialize_statement_contexts(self) -> dict:
        """Initialize statement contexts with mapping for Arelle GUI compatibility"""
        statement_contexts = {
            'BS': {},
            'IS': {},
            'CI': {},
            'CF': {},
            'EQ': {}
        }
        
        log_debug("\n=== Initializing Statement Contexts for Arelle GUI ===")
        
        try:
            relationship_set = self.model_xbrl.relationshipSet(XbrlConst.parentChild)            
            if not relationship_set:
                log_debug("No relationship set found")
                return statement_contexts

            # Track role types and their facts for Arelle GUI
            for roleURI in relationship_set.linkRoleUris:
                roleTypes = self.model_xbrl.roleTypes.get(roleURI, [])
                if roleTypes:
                    roleType = roleTypes[0]
                    stmt_type = getattr(roleType, '_statementType', '').rstrip('P')
                    
                    if stmt_type in statement_contexts:
                        facts = getattr(roleType, '_tableFacts', set())
                        for fact in facts:
                            if fact.context is None:
                                continue
                                
                            context = fact.context
                            context_id = context.id
                            
                            # Store context with its period info
                            if context.isInstantPeriod:
                                period = (None, context.instantDatetime)
                            else:
                                period = (context.startDatetime, context.endDatetime)
                                
                            if period[1]:  # Only store if we have an end date
                                statement_contexts[stmt_type][context_id] = period
                                
                            # Set properties needed by Arelle GUI
                            roleType._tableContexts = {context_id: period}
                            if not hasattr(roleType, '_viewProperties'):
                                roleType._viewProperties = {}
                            roleType._viewProperties['periods'] = [period]
                            
            return statement_contexts
                    
        except Exception as e:
            log_debug(f"Error initializing contexts: {str(e)}", include_trace=True)
            return statement_contexts

    def get_role_info(self, definition: str) -> dict:
        """Initial categorization based on basic role patterns"""
        result = {
            'is_statement': bool(re.match(self.ROLE_PATTERNS['STMT'], definition.lower())),
            'is_document': bool(re.match(self.ROLE_PATTERNS['DOC'], definition.lower())),
            'is_disclosure': bool(re.match(self.ROLE_PATTERNS['DISCLOSURE'], definition.lower())),
            'has_details': not bool(re.match(self.ROLE_PATTERNS['DETAILS'], definition.lower())),
            'is_comprehensive': bool(re.match(self.ROLE_PATTERNS['IS_COMPREHENSIVE'], definition.lower())),
            'is_parenthetical': bool(re.search(self.ROLE_PATTERNS['PARENTHETICAL'], definition.lower()))
        }
        return result



class TableDataTransformer:
    """Transforms table data to show periods as columns while preserving original order"""
    
    def __init__(self):
        self.logger = logger
        
    def _extract_period_key(self, row):
        """Creates a period key from start and end dates"""
        start_date = row[5]  # Start Date column
        end_date = row[6]    # End Date column
        
        if end_date and not start_date:
            return f"As of {end_date}"  # For instant periods
        elif start_date and end_date:
            return f"{start_date} - {end_date}"  # For duration periods
        return "Unknown Period"
        
    def _get_unique_periods(self, data):
        """Extract and sort all unique periods from the data"""
        periods = set()
        for row in data:
            periods.add(self._extract_period_key(row))
        return sorted(list(periods))
        
    def transform_table_data(self, headers, data):
        """
        Transform table data to show periods as columns while maintaining original order
        
        Args:
            headers: Original column headers
            data: Original data rows
            
        Returns:
            new_headers: New column headers with periods
            new_data: Transformed data with one row per concept
        """
        try:
            # Get sorted list of all periods first
            unique_periods = self._get_unique_periods(data)
            
            # Create new headers
            new_headers = ['Concept', 'QName', 'Unit']
            new_headers.extend(unique_periods)
            
            # Group data by concept while preserving order
            ordered_concepts = []
            concept_data = {}
            seen_concepts = set()
            
            for row in data:
                label = row[0]  # Concept column
                qname = row[1]  # QName column
                value = row[2]  # Value column
                unit = row[3]   # Unit column
                
                concept_key = (qname, label)
                if concept_key not in seen_concepts:
                    ordered_concepts.append(concept_key)
                    seen_concepts.add(concept_key)
                    concept_data[concept_key] = {
                        'label': label,
                        'qname': qname,
                        'unit': unit,
                        'periods': {}
                    }
                
                period_key = self._extract_period_key(row)
                concept_data[concept_key]['periods'][period_key] = value
            
            # Create new data rows maintaining original order
            new_data = []
            for concept_key in ordered_concepts:
                concept_info = concept_data[concept_key]
                row = [
                    concept_info['label'],
                    concept_info['qname'],
                    concept_info['unit']
                ]
                
                # Add values for each period
                for period in unique_periods:
                    row.append(concept_info['periods'].get(period, ''))
                    
                new_data.append(row)
                
            return new_headers, new_data
            
        except Exception as e:
            self.logger.log(f"Error transforming table data: {str(e)}", "ERROR", True)
            return headers, data  # Return original data on error



class AccessTracker:
    """Tracks and validates value access patterns"""
    def __init__(self):
        self.access_patterns = defaultdict(list)
        self.conflicts = []
        self.logger = logger

    def track_access(self, concept: str, access_type: str, metadata: dict):
        """Track each value access attempt"""
        self.access_patterns[concept].append({
            'type': access_type,  # 'direct', 'cached', 'managed'
            'timestamp': datetime.now(),
            'metadata': metadata
        })
        self._check_for_conflicts(concept)

    def _check_for_conflicts(self, concept: str):
        """Check for conflicting access patterns"""
        recent_accesses = self.access_patterns[concept][-5:]  # Look at last 5 accesses
        if len(recent_accesses) < 2:
            return

        has_direct = any(a['type'] == 'direct' for a in recent_accesses)
        has_managed = any(a['type'] == 'managed' for a in recent_accesses)

        if has_direct and has_managed:
            conflict = {
                'concept': concept,
                'timestamp': datetime.now(),
                'message': f"Mixed access patterns detected for {concept}",
                'patterns': recent_accesses
            }
            self.conflicts.append(conflict)
            log_debug(f"\nWARNING: Access pattern conflict detected:")
            log_debug(f"Concept: {concept}")
            log_debug("Recent accesses:")
            for access in recent_accesses:
                log_debug(f"- Type: {access['type']}, Time: {access['timestamp']}")

class ExcelFormatter:
    """Handles Excel formatting and sheet structure using openpyxl"""
    
    def __init__(self, writer):
        self.writer = writer
        self.workbook = writer.book
        
    def format_statement_sheet(self, sheet_name: str, df: pd.DataFrame):
        """Format a financial statement sheet"""
        worksheet = self.writer.sheets[sheet_name]
        
        # Define styles using openpyxl
        header_style = {
            'fill': PatternFill(fill_type='solid', start_color='D9E1F2'),
            'font': Font(bold=True),
            'border': Border(
                left=Side(style='thin'),
                right=Side(style='thin'),
                top=Side(style='thin'),
                bottom=Side(style='thin')
            )
        }
        
        number_style = {
            'number_format': '#,##0.00',
            'border': Border(
                left=Side(style='thin'),
                right=Side(style='thin'),
                top=Side(style='thin'),
                bottom=Side(style='thin')
            )
        }
        
        text_style = {
            'border': Border(
                left=Side(style='thin'),
                right=Side(style='thin'),
                top=Side(style='thin'),
                bottom=Side(style='thin')
            )
        }
        
        # Apply styles
        for idx, col in enumerate(df.columns):
            cell = worksheet.cell(row=1, column=idx + 1)
            cell.value = col
            # Apply header styles
            cell.fill = header_style['fill']
            cell.font = header_style['font']
            cell.border = header_style['border']
            
            # Set column width
            max_length = max(
                df[col].astype(str).apply(len).max(),
                len(str(col))
            )
            worksheet.column_dimensions[get_column_letter(idx + 1)].width = max_length + 2
            
            # Apply styles to data cells
            is_number_col = 'Value' in col or any(w in col for w in ['As of', 'to'])
            for row_idx in range(len(df)):
                cell = worksheet.cell(row=row_idx + 2, column=idx + 1)
                if is_number_col:
                    if isinstance(cell.value, (int, float)):
                        cell.number_format = number_style['number_format']
                cell.border = text_style['border']

    def format_ratio_details_sheet(self, df: pd.DataFrame, sheet_name: str = 'Ratio_Details'):
        """Format the ratio details sheet with enhanced context awareness"""
        worksheet = self.writer.sheets[sheet_name]
        
        # Define styles
        header_style = {
            'fill': PatternFill(fill_type='solid', start_color='E2EFDA'),
            'font': Font(bold=True),
            'border': Border(
                left=Side(style='thin'),
                right=Side(style='thin'),
                top=Side(style='thin'),
                bottom=Side(style='thin')
            )
        }
        
        context_style = {
            'fill': PatternFill(fill_type='solid', start_color='F2F2F2'),
            'font': Font(italic=True),
            'number_format': '@'  # Text format for context IDs
        }
        
        value_style = {
            'number_format': '#,##0.00'
        }
        
        # Apply styles with context awareness
        for idx, col in enumerate(df.columns):
            cell = worksheet.cell(row=1, column=idx + 1)
            cell.value = col
            cell.fill = header_style['fill']
            cell.font = header_style['font']
            cell.border = header_style['border']
            
            # Set column formatting based on content type
            is_context_col = 'Context' in col
            is_value_col = 'Value' in col or any(w in col for w in ['Ratio', 'Margin', 'Return'])
            
            for row_idx in range(len(df)):
                cell = worksheet.cell(row=row_idx + 2, column=idx + 1)
                
                if is_context_col:
                    cell.fill = context_style['fill']
                    cell.font = context_style['font']
                    cell.number_format = context_style['number_format']
                elif is_value_col and isinstance(cell.value, (int, float)):
                    cell.number_format = value_style['number_format']
            
            # Adjust column width
            max_length = max(
                df[col].astype(str).apply(len).max(),
                len(str(col))
            )
            worksheet.column_dimensions[get_column_letter(idx + 1)].width = max_length + 2

    def format_financial_ratios_sheet(self, df: pd.DataFrame):
        """Format the new Financial_Ratios sheet with period-based columns"""
        worksheet = self.writer.sheets['Financial_Ratios']
        
        # Define styles
        header_style = {
            'fill': PatternFill(fill_type='solid', start_color='E2EFDA'),
            'font': Font(bold=True),
            'border': Border(
                left=Side(style='thin'),
                right=Side(style='thin'),
                top=Side(style='thin'),
                bottom=Side(style='thin')
            )
        }
        
        value_style = {
            'number_format': '#,##0.00',
            'alignment': Alignment(horizontal='right')
        }
        
        text_style = {
            'alignment': Alignment(horizontal='left')
        }
        
        # Apply styles
        for idx, col in enumerate(df.columns):
            cell = worksheet.cell(row=1, column=idx + 1)
            cell.value = col
            cell.fill = header_style['fill']
            cell.font = header_style['font']
            cell.border = header_style['border']
            
            # Apply appropriate formatting based on column type
            is_value_col = idx > 1  # After Ratio Name and Statement Types
            
            for row_idx in range(len(df)):
                cell = worksheet.cell(row=row_idx + 2, column=idx + 1)
                
                if is_value_col:
                    if isinstance(cell.value, (int, float)):
                        cell.number_format = value_style['number_format']
                        cell.alignment = value_style['alignment']
                else:
                    cell.alignment = text_style['alignment']
            
            # Adjust column width
            max_length = max(
                df[col].astype(str).apply(len).max(),
                len(str(col))
            )
            worksheet.column_dimensions[get_column_letter(idx + 1)].width = max_length + 2

    def format_inputs_sheet(self, df: pd.DataFrame):
        """Format the Ratio_Inputs sheet with enhanced grouping"""
        worksheet = self.writer.sheets['Ratio_Inputs']
        
        # Define styles
        header_style = {
            'fill': PatternFill(fill_type='solid', start_color='FCE4D6'),
            'font': Font(bold=True),
            'border': Border(
                left=Side(style='thin'),
                right=Side(style='thin'),
                top=Side(style='thin'),
                bottom=Side(style='thin')
            )
        }
        
        group_style = {
            'fill': PatternFill(fill_type='solid', start_color='F2F2F2'),
            'font': Font(bold=True),
            'border': Border(
                left=Side(style='thin'),
                right=Side(style='thin'),
                top=Side(style='thin'),
                bottom=Side(style='thin')
            )
        }
        
        value_style = {
            'number_format': '#,##0.00'
        }
        
        # Apply styles
        for idx, col in enumerate(df.columns):
            cell = worksheet.cell(row=1, column=idx + 1)
            cell.value = col
            cell.fill = header_style['fill']
            cell.font = header_style['font']
            cell.border = header_style['border']
            
            # Apply number format to value columns (typically after column 2)
            is_value_col = idx > 2
            
            for row_idx in range(len(df)):
                cell = worksheet.cell(row=row_idx + 2, column=idx + 1)
                
                # Apply grouping style to group headers
                if isinstance(cell.value, str) and not cell.value.startswith('  '):
                    cell.fill = group_style['fill']
                    cell.font = group_style['font']
                
                # Apply number format to values
                if is_value_col and isinstance(cell.value, (int, float)):
                    cell.number_format = value_style['number_format']
            
            # Adjust column width
            max_length = max(
                df[col].astype(str).apply(len).max(),
                len(str(col))
            )
            worksheet.column_dimensions[get_column_letter(idx + 1)].width = max_length + 2

    def format_disclaimer_sheet(self, worksheet):
        """Format the disclaimer sheet with corporate styling"""
        # Define styles
        title_style = {
            'font': Font(bold=True, size=14),
            'alignment': Alignment(horizontal='center', vertical='center')
        }
        
        text_style = {
            'font': Font(size=11),
            'alignment': Alignment(wrap_text=True, vertical='top')
        }
        
        # Set column width
        worksheet.column_dimensions['A'].width = 100
        
        # Add and format title
        title_cell = worksheet.cell(row=1, column=1)
        title_cell.value = "Disclaimer"
        title_cell.font = title_style['font']
        title_cell.alignment = title_style['alignment']
        
        # Add disclaimer text
        disclaimer_text = (
            "Financial Analysis Disclaimer\n\n"
            "The financial ratios and analysis presented in this report are calculated based on the data "
            "extracted from XBRL filings. Please note the following important considerations:\n\n"
            "1. The calculations are automated and should be verified against the original financial statements.\n\n"
            "2. Different accounting standards and company-specific reporting practices may affect the "
            "comparability of ratios across different companies or periods.\n\n"
            "3. This analysis is provided for informational purposes only and should not be considered as "
            "financial advice or a recommendation for any investment decision.\n\n"
            "4. Users should conduct their own due diligence and consult with qualified financial advisors "
            "before making any financial decisions.\n\n"
            "5. The accuracy of these ratios depends on the accuracy and completeness of the underlying "
            "XBRL data in the filing.\n\n"
            "6. The creators and contributors of this software plugin are not liable for any errors, "
            "omissions, or any consequences arising from the use of this analysis tool. Use of this "
            "software is at your own risk."            
        )
        
        # Add text with proper formatting
        text_cell = worksheet.cell(row=3, column=1)
        text_cell.value = disclaimer_text
        text_cell.font = text_style['font']
        text_cell.alignment = text_style['alignment']
        
        # Set row heights
        worksheet.row_dimensions[1].height = 30  # Title row
        worksheet.row_dimensions[2].height = 20  # Spacing


class ExcelExporter:
    """Enhanced Excel exporter with consolidated data handling"""
    
    def __init__(self, model_xbrl):
        self.model_xbrl = model_xbrl
        # self.output_dir = r"C:\Data\venv_ArellePlugin_Release\sec_tables_output"  # hardcoded path during development
        # os.makedirs(self.output_dir, exist_ok=True)
        self.fact_provider = model_xbrl._fact_provider
        self.data_manager = self.fact_provider.financial_data_manager
        self.ratio_calculator = FinancialRatioCalculator(model_xbrl)
        log_debug("\nInitializing Excel Exporter with FinancialRatioCalculator (ExcelExporter.__init__)")  # Add only this line
        self.logger = logger
        self.transformer = TableDataTransformer()

    def prepare_ratio_inputs(self) -> dict:
        """Prepare ratio input data using context mappings"""
        log_debug("\n=== Preparing Ratio Inputs Data (ExcelExporter.prepare_ratio_inputs) ===")
        
        # Get valid context pairs
        context_pairs = self._get_valid_context_pairs()
        if not context_pairs:
            log_debug("No valid context pairs found (ExcelExporter.prepare_ratio_inputs)")
            return {
                'headers': ['Input', 'Origin', 'QName'],
                'data': []
            }
                
        # Sort context pairs by end date
        def get_end_date(ctx_pair):
            bs_ctx, is_ctx = ctx_pair
            bs_period = self.data_manager.get_period_for_context(bs_ctx, 'BS')
            if bs_period and bs_period[1]:
                return bs_period[1]
            is_period = self.data_manager.get_period_for_context(is_ctx, 'IS')
            return is_period[1] if is_period else datetime.min
        
        sorted_pairs = sorted(context_pairs, key=get_end_date)
        
        # Log the sorted pairs
        log_debug("\nProcessing sorted context pairs:")
        for bs_ctx, is_ctx in sorted_pairs:
            bs_period = self.data_manager.get_period_for_context(bs_ctx, 'BS')
            is_period = self.data_manager.get_period_for_context(is_ctx, 'IS')
            log_debug(f"\nPair:")
            if bs_period and bs_period[1]:
                log_debug(f"  BS: {bs_ctx} - End date: {bs_period[1].strftime('%Y-%m-%d')}")
            if is_period and is_period[1]:
                log_debug(f"  IS: {is_ctx} - End date: {is_period[1].strftime('%Y-%m-%d')}")
        
        # Get inputs data through ratio calculator
        log_debug(f"\nProcessing {len(sorted_pairs)} context pairs")
        return self.ratio_calculator.get_inputs_data(sorted_pairs)
 
    def _log_export_results(self, successful_exports: List[tuple], failed_exports: List[tuple]):
        """Log export results"""
        log_debug(f"\nSuccessfully exported {len(successful_exports)} sheets:")
        for sheet_name, full_name in successful_exports:
            log_debug(f"- {sheet_name}: {full_name}")

        if failed_exports:
            log_debug("\nFailed exports:")
            for name, reason in failed_exports:
                log_debug(f"- {name}: {reason}")

    def get_sheet_name(self, statement_type: str, seq_num: str) -> str:
        """Generate Excel-compatible sheet name"""
        # Use statement type and sequence number
        sheet_name = f"{seq_num}-{statement_type}"
        
        # Clean up sheet name
        sheet_name = re.sub(r'[\\/:*?\[\]]', '', sheet_name)
        sheet_name = re.sub(r'\s+', ' ', sheet_name)
        sheet_name = sheet_name.strip()
        
        # Excel limits sheet names to 31 characters
        if len(sheet_name) > 31:
            sheet_name = sheet_name[:31]
            
        return sheet_name

    def get_financial_statements(self):
        """Get all financial statements using enhanced detection"""
        log_debug("\nCollecting financial statements for export...")
        tables_data = []
        found_statements = set()
        
        try:
            # Get all role types using TableStructure's order
            relationshipSet = self.model_xbrl.relationshipSet(XbrlConst.parentChild)
            if not relationshipSet:
                log_debug("No relationship set found")
                return []
                
            # Get sorted role types with their statement information
            sorted_role_types = []
            for roleURI in relationshipSet.linkRoleUris:
                roleTypes = self.model_xbrl.roleTypes.get(roleURI)
                if roleTypes:
                    roleType = roleTypes[0]
                    definition = (roleType.genLabel(strip=True) or 
                                roleType.definition or 
                                roleURI)
                    # Use enhanced statement detection info
                    tableIndex = getattr(roleType, '_tableIndex', None)
                    statementType = getattr(roleType, '_statementType', None)
                    if tableIndex:
                        sorted_role_types.append((definition, roleType, tableIndex, statementType))
            
            # Sort using TableStructure's assigned group and sequence
            sorted_role_types.sort(key=lambda rt: (rt[2][0], rt[2][1]) if rt[2] else ("", ""))
            
            log_debug("\nProcessing statements in order:")
            for roleDefinition, roleType, tableIndex, statementType in sorted_role_types:
                try:
                    log_debug(f"\nProcessing: {roleDefinition}")
                    log_debug(f"Statement Type: {statementType}")
                    
                    # Only process financial statements
                    if not tableIndex or tableIndex[0] != "2Financial Statements":
                        log_debug("Skipping - not a financial statement")
                        continue
                        
                    tableFacts = getattr(roleType, '_tableFacts', None)
                    if not tableFacts:
                        log_debug("No table facts found")
                        continue
                        
                    # Get relationship set for this role
                    roleRelationshipSet = self.model_xbrl.relationshipSet(XbrlConst.parentChild, 
                                                                        roleType.roleURI)
                    
                    if not roleRelationshipSet:
                        log_debug("No relationship set for role")
                        continue
                        
                    # Process the statement
                    headers, data = self.process_statement(roleType, tableFacts, roleRelationshipSet)
                    
                    if data:
                        statement_info = {
                            'headers': headers,
                            'data': data,
                            'statement_name': roleDefinition,
                            'statement_type': statementType or 'OTH',
                            'table_index': tableIndex,
                            'sequence': tableIndex[1]
                        }
                        
                        found_statements.add(f"{roleDefinition}  - Type: {statementType or 'Unknown'}")
                        tables_data.append(statement_info)
                        log_debug(f"Added statement to export list: {roleDefinition}")
                    else:
                        log_debug("No data processed")
                        
                except Exception as e:
                    log_debug(f"Error processing role type: {str(e)}")
                    log_debug(traceback.format_exc())
                    continue
            
            # Write summary
            log_debug("\nFinancial Statements Found:")
            for stmt in sorted(found_statements):
                log_debug(f"- {stmt}")
                
            return tables_data
                
        except Exception as e:
            log_debug(f"Error in get_financial_statements: {str(e)}")
            log_debug(traceback.format_exc())
            return []

    def merge_split_statements(self, statements):
        """
        Merge split statements (e.g., Aktiva/Passiva) into single statements
        with enhanced pattern matching and debugging
        """
        log_debug("\n=== Starting Split Statement Merge Process ===")
        merged = []
        pending_merge = {}
        
        # Enhanced split detection patterns
        split_patterns = {
            'assets': [
                r'(?i).*aktiva.*',
                r'(?i).*assets?.*',
                r'(?i).*aktivseite.*',
                r'(?i).*statement.*assets.*',
                r'(?i).*consolidated.*assets.*'
            ],
            'liabilities': [
                r'(?i).*passiva.*',
                r'(?i).*liabilit.*',
                r'(?i).*passivseite.*',
                r'(?i).*equity.*',
                r'(?i).*statement.*liab.*'
            ]
        }

        for stmt in statements:
            stmt_name = stmt['statement_name'].lower()
            stmt_type = stmt['statement_type']
            log_debug(f"\nProcessing statement: {stmt['statement_name']} (Type: {stmt_type})")
            
            # Check if this is a balance sheet statement that might be split
            if stmt_type == 'BS':
                # Enhanced split detection
                is_assets = any(re.search(pattern, stmt_name) for pattern in split_patterns['assets'])
                is_liabilities = any(re.search(pattern, stmt_name) for pattern in split_patterns['liabilities'])
                
                log_debug(f"Split detection results:")
                log_debug(f"  Assets patterns: {is_assets}")
                log_debug(f"  Liabilities patterns: {is_liabilities}")
                
                if is_assets or is_liabilities:
                    section = 'assets' if is_assets else 'liabilities'
                    log_debug(f"Detected split BS section: {section}")
                    
                    # Generate consistent key for matching splits
                    key = f"BS_{stmt.get('sequence', '00')}"
                    log_debug(f"Using merge key: {key}")
                    
                    if key in pending_merge:
                        log_debug("Found matching section for merge")
                        # Merge the data
                        merged_data = pending_merge[key]['data']
                        original_count = len(merged_data)
                        
                        # Ensure proper order (assets before liabilities)
                        if section == 'assets':
                            new_data = stmt['data'] + merged_data
                        else:
                            new_data = merged_data + stmt['data']
                        
                        # Create merged statement
                        merged_stmt = pending_merge[key].copy()
                        merged_stmt['data'] = new_data
                        merged_stmt['statement_name'] = "Balance Sheet"  # Generic name
                        
                        log_debug(f"Merged {original_count} + {len(stmt['data'])} = {len(new_data)} rows")
                        
                        # Add period information if available
                        if 'period' in stmt:
                            merged_stmt['period'] = stmt['period']
                        
                        merged.append(merged_stmt)
                        del pending_merge[key]
                        log_debug("Merge completed")
                    else:
                        log_debug("Storing for later merge")
                        pending_merge[key] = stmt
                else:
                    log_debug("Not a split BS - adding directly")
                    merged.append(stmt)
            else:
                log_debug("Not a BS statement - adding directly")
                merged.append(stmt)
        
        # Handle any unmerged splits
        if pending_merge:
            log_debug("\nWARNING: Some split statements could not be merged:")
            for key, stmt in pending_merge.items():
                log_debug(f"- Unmerged section: {stmt['statement_name']}")
                # Add unmerged statements to output
                merged.append(stmt)
        
        # Log final results
        log_debug(f"\nMerge process complete. Final statement count: {len(merged)}")
        log_debug("Final statements:")
        for stmt in merged:
            log_debug(f"- {stmt['statement_name']} ({len(stmt['data'])} rows)")
        
        return merged

    def process_statement(self, roleType, facts, relationshipSet) -> Tuple[List, List]:
        """Process statement with enhanced data flow"""
        log_debug(f"\nProcessing statement: {roleType.definition}")
        
        # Get valid context periods
        context_periods = self._get_valid_periods(facts)
        log_debug(f"Found {len(context_periods)} valid periods")
        
        # Get filtered facts through data manager
        filtered_facts = self._get_filtered_facts(facts, context_periods)
        log_debug(f"Filtered to {len(filtered_facts)} facts")
        
        # Process concepts and get data
        headers, data = self._process_concepts(roleType, filtered_facts, relationshipSet)
        
        # Transform data to period columns
        transformed_headers, transformed_data = self.transformer.transform_table_data(headers, data)
        
        return transformed_headers, transformed_data
        
    def _get_valid_periods(self, facts) -> Set[tuple]:
        """Get valid periods from facts using data manager"""
        periods = set()
        for fact in facts:
            if not fact.context:
                continue
            
            period_key = self.data_manager.get_period_key(fact.context)
            if period_key:
                periods.add(period_key)
        return periods
        
    def _get_filtered_facts(self, facts, valid_periods) -> List:
        """Filter facts based on valid periods"""
        filtered = []
        for fact in facts:
            if not fact.context:
                continue
                
            period_key = self.data_manager.get_period_key(fact.context)
            if period_key in valid_periods:
                filtered.append(fact)
        return filtered
        
    def _process_concepts(self, roleType, facts, relationshipSet) -> Tuple[List, List]:
        """Process concepts with data manager integration"""
        headers = ['Concept', 'QName', 'Value', 'Unit', 'Decimals', 'Start Date', 'End Date', 'EntityID']
        data = []
        
        for rootConcept in relationshipSet.rootConcepts:
            data.extend(self._process_concept_facts(rootConcept, facts, relationshipSet))
            
        return headers, data
        
    def _process_concept_facts(self, concept, facts, relationshipSet, indent_level=0) -> List[List]:
        """Process concept facts with enhanced metadata tracking"""
        data = []
        
        # Get preferred label
        preferredLabel = self._get_preferred_label(concept, relationshipSet)
        
        # Group facts by context using data manager
        facts_by_context = self._group_facts_by_context(concept, facts)
        
        # Process each context group
        for period_key, period_facts in facts_by_context.items():
            for fact in period_facts:
                data.append(self._create_fact_row(fact, concept, preferredLabel, indent_level))
                
        # Process child concepts
        for rel in relationshipSet.fromModelObject(concept):
            if rel.toModelObject is not None:
                data.extend(self._process_concept_facts(
                    rel.toModelObject, facts, relationshipSet, indent_level + 1))
                
        return data
        
    def _get_preferred_label(self, concept, relationshipSet) -> Optional[str]:
        """Get preferred label with error handling"""
        try:
            for rel in relationshipSet.toModelObject(concept):
                if rel.preferredLabel:
                    return rel.preferredLabel
        except Exception as e:
            self.logger.log(f"Error getting preferred label: {str(e)}")
        return None
        
    def _group_facts_by_context(self, concept, facts) -> Dict[tuple, List]:
        """Group facts by context period"""
        grouped = defaultdict(list)
        concept_facts = [f for f in facts if f.concept == concept]
        
        for fact in concept_facts:
            if fact.context:
                period_key = self.data_manager.get_period_key(fact.context)
                if period_key:
                    grouped[period_key].append(fact)
                    
        return grouped
        
    def _create_fact_row(self, fact, concept, preferredLabel, indent_level) -> List:
        """Create a fact data row with proper formatting"""
        indent = "  " * indent_level
        start_date, end_date = self.data_manager.get_period_dates(fact.context)
        
        return [
            f"{indent}{concept.label(preferredLabel)}",
            str(concept.qname),
            fact.value if fact.value is not None else "",
            str(fact.unit) if fact.unit else "",
            str(fact.decimals) if fact.decimals is not None else "",
            start_date.strftime('%Y-%m-%d') if start_date else "",
            end_date.strftime('%Y-%m-%d') if end_date else "",
            fact.context.entityIdentifier[1] if fact.context else ""
        ]

    def prepare_excel_data(self) -> dict:
        """Prepare all data for Excel export using FinancialRatioCalculator"""        
        log_debug("\n=== Preparing Excel Data (ExcelExporter.prepare_excel_data) ===")
        
        empty_structure = {
            'statements': [],
            'ratios': {'headers': [], 'data': []},
            'inputs': {'headers': [], 'data': []}
        }   
                    
        try:
            # Get all data first
            statements = self.get_financial_statements()
            if not statements:
                log_debug("No statements found - returning empty structure")
                return empty_structure
                
            merged_statements = self.merge_split_statements(statements) if statements else []
            
            # Let FinancialRatioCalculator handle all ratio calculations
            ratio_data = self.ratio_calculator.prepare_ratio_data()
            inputs_data = self.prepare_ratio_inputs()
            
            result = {
                'statements': merged_statements,
                'ratios': ratio_data,
                'inputs': inputs_data
            }
            
            log_debug("\nData summary (ExcelExporter.prepare_excel_data):")
            log_debug(f"- Statements: {len(result['statements'])}")
            log_debug(f"- Ratio rows: {len(result['ratios'].get('data', []))}")
            log_debug(f"- Input rows: {len(result['inputs'].get('data', []))}")
            
            return result
                
        except Exception as e:
            log_debug(f"Error preparing Excel data: {str(e)}", include_trace=True)
            return empty_structure

    def export_tables_to_excel(self):
        """Export with enhanced ratio handling and contextRef support"""
        try:
            filename = os.path.basename(self.model_xbrl.uri)
            base_filename = os.path.splitext(filename)[0]
            
            # old hardcoded
            # output_path = os.path.join(self.output_dir, f"{base_filename}_financial_statements.xlsx") # Hardcoded
            # new config file based
            export_dir = self.model_xbrl._config_handler.get_export_path(self.model_xbrl)
            if not export_dir:
                print("Excel export disabled in config")
                log_debug("Excel export disabled in config")
                return False
                
            os.makedirs(export_dir, exist_ok=True)
            output_path = os.path.join(export_dir, f"{base_filename}_financial_statements.xlsx")   
            print(f"\nWriting Excel file to: {output_path}")
            log_debug(f"\nWriting Excel file to: {output_path}")                  
            
            
            
            log_debug("\n=== Starting Excel Export (ExcelExporter.export_tables_to_excel)===")
            
            # Get all data first
            excel_data = self.prepare_excel_data()
            
            
            # Log detailed data availability
            log_debug("\nChecking data availability:")
            log_debug(f"- Statements available: {bool(excel_data.get('statements', []))}")
            if excel_data.get('statements'):
                for stmt in excel_data['statements']:
                    log_debug(f"  * {stmt.get('statement_name', 'Unknown')}: {len(stmt.get('data', []))} rows")
                    
            log_debug(f"- Ratios available: {bool(excel_data.get('ratios', {}).get('data', []))}")
            if excel_data.get('ratios', {}).get('data'):
                log_debug(f"  * Ratio rows: {len(excel_data['ratios']['data'])}")
                
            log_debug(f"- Inputs available: {bool(excel_data.get('inputs', {}).get('data', []))}")
            if excel_data.get('inputs', {}).get('data'):
                log_debug(f"  * Input rows: {len(excel_data['inputs']['data'])}")
                            
            # Now safe to check if we have any data to export
            has_statements = bool(excel_data.get('statements', []))
            has_ratios = bool(excel_data.get('ratios', {}).get('data', []))
            has_inputs = bool(excel_data.get('inputs', {}).get('data', []))
            
            if not any([has_statements, has_ratios, has_inputs]):
                log_debug("No data available for export (ExcelExporter.export_tables_to_excel)")
                return False
            
            successful_exports = []
            failed_exports = []
            
            with pd.ExcelWriter(output_path, engine='openpyxl') as writer:
                formatter = ExcelFormatter(writer)

                # Create Disclaimer sheet first
                workbook = writer.book
                disclaimer_sheet = workbook.create_sheet("Disclaimer")
                formatter.format_disclaimer_sheet(disclaimer_sheet)
                successful_exports.append(('Disclaimer', 'Disclaimer'))                
                
                # Export statements if we have them - this should work independently of ratios
                if excel_data['statements']:
                    for stmt in excel_data['statements']:
                        try:
                            sheet_name = self.get_sheet_name(stmt['statement_type'], stmt['sequence'])
                            df = pd.DataFrame(stmt['data'], columns=stmt['headers'])
                            if not df.empty:
                                df.to_excel(writer, sheet_name=sheet_name, index=False)
                                formatter.format_statement_sheet(sheet_name, df)
                                successful_exports.append((sheet_name, stmt['statement_name']))
                                log_debug(f"Exported statement: {sheet_name}")
                        except Exception as e:
                            failed_exports.append((stmt['statement_name'], str(e)))
                            log_debug(f"Error exporting statement: {str(e)}")


                # Try to export Ratio_Details and Financial_Ratios if available
                if excel_data['ratios']['data']:
                    try:
                        # Export new columnar ratio sheet
                        transformed_data = self.transform_ratio_data_for_columns(excel_data['ratios'])
                        if transformed_data['data']:
                            df_ratios = pd.DataFrame(transformed_data['data'],
                                                columns=transformed_data['headers'])
                            df_ratios.to_excel(writer, sheet_name='Financial_Ratios', index=False)
                            formatter.format_financial_ratios_sheet(df_ratios)
                            successful_exports.append(('Financial_Ratios', 'Financial Ratios'))
                            log_debug("Successfully exported both ratio sheets")
                        else:
                            log_debug("No data available for Financial_Ratios sheet")
                            
                            
                        # Export detailed ratio sheet
                        df_details = pd.DataFrame(excel_data['ratios']['data'], 
                                                columns=excel_data['ratios']['headers'])
                        df_details.to_excel(writer, sheet_name='Ratio_Details', index=False)
                        formatter.format_ratio_details_sheet(df_details)
                        successful_exports.append(('Ratio_Details', 'Ratio Details'))
                                                    
                    except Exception as e:
                        log_debug(f"Error exporting ratio sheets: {str(e)}", include_trace=True)
                        failed_exports.append(('Ratio_Sheets', str(e)))
                                        
                
                # Ratio_Inputs with Excessive Logging
                log_debug("\nProcessing ratio inputs data:")
                log_debug(f"Input data structure: {excel_data.get('inputs', {}).keys()}")
                if excel_data.get('inputs', {}).get('data'):
                    try:
                        input_data = excel_data['inputs']['data']
                        input_headers = excel_data['inputs']['headers']
                        log_debug(f"Input data rows: {len(input_data)}")
                        log_debug(f"Input headers: {input_headers}")
                        
                        df = pd.DataFrame(input_data, columns=input_headers)
                        log_debug(f"Created DataFrame with shape: {df.shape}")
                        
                        df.to_excel(writer, sheet_name='Ratio_Inputs', index=False)
                        log_debug("Written to Excel sheet")
                        
                        formatter.format_inputs_sheet(df)
                        log_debug("Applied formatting")
                        
                        successful_exports.append(('Ratio_Inputs', 'Ratio Inputs'))
                        log_debug("Successfully added Ratio_Inputs sheet")
                    except Exception as e:
                        log_debug(f"Error creating Ratio_Inputs sheet: {str(e)}", include_trace=True)
                        failed_exports.append(('Ratio_Inputs', str(e)))
                else:
                    log_debug("No input data available for Ratio_Inputs sheet")
                    log_debug(f"excel_data structure: {excel_data.keys()}")
                    log_debug(f"inputs structure: {excel_data.get('inputs', {})}")                
                
                # Only create empty sheet if absolutely nothing was exported
                if not successful_exports:
                    pd.DataFrame({'No Data': []}).to_excel(writer, sheet_name='No_Data', index=False)
                    successful_exports.append(('No_Data', 'Empty Sheet'))
                
                self._log_export_results(successful_exports, failed_exports)
                
            log_debug("Excel export completed successfully")
            print(f"Successfully wrote Excel file: {output_path}")
            log_debug(f"Successfully wrote Excel file: {output_path}")            
            return True
            
        except Exception as e:
            log_debug(f"Error during export: {str(e)}", include_trace=True)
            return False
            
    def prepare_ratio_inputs(self) -> dict:
        """Prepare ratio input data using context mappings"""
        log_debug("\nPreparing ratio input data (ExcelExporter.prepare_ratio_inputs)")
        
        # Get all context pairs
        context_pairs = self._get_valid_context_pairs()
        if not context_pairs:
            log_debug("No valid context pairs found (ExcelExporter.prepare_ratio_inputs)")
            return {
                'headers': ['Input', 'Origin', 'QName'],
                'data': []
            }
            
        return self.ratio_calculator.get_inputs_data(context_pairs)    

    def prepare_ratio_data(self) -> dict:
        """Prepare ratio data with context-based grouping"""
        log_debug("\n=== Preparing Ratio Data (ExcelExporter.prepare_ratio_data) ===")
        
        try:
            # Get valid context pairs
            context_pairs = self._get_valid_context_pairs()
            if not context_pairs:
                log_debug("No valid context pairs found (ExcelExporter.prepare_ratio_data)")
                return {'headers': [], 'data': []}
                    
            # Calculate ratios for each context set
            for bs_ctx, is_ctx in context_pairs:
                log_debug(f"\nProcessing context set for BS: {bs_ctx}, IS: {is_ctx}")
                
                # Get CF context BEFORE calculating ratios
                cf_contexts = self.data_manager.find_matching_cf_contexts(bs_ctx)
                cf_ctx = cf_contexts[0] if cf_contexts else None
                log_debug(f"Found CF context: {cf_ctx}")

                if cf_contexts:
                    cf_period = self.data_manager.get_period_for_context(cf_ctx, 'CF')
                    if cf_period:
                        log_debug(f"CF Period: {cf_period[0].strftime('%Y-%m-%d') if cf_period[0] else 'None'} to {cf_period[1].strftime('%Y-%m-%d')}")
                
                result_set = self.ratio_calculator.calculate_ratios_for_contexts(bs_ctx, is_ctx, cf_ctx)
                if result_set and result_set.results:
                    log_debug(f"Calculated {len(result_set.results)} ratios")
                    self.ratio_analysis.add_result_set(result_set)
                        
            # Get formatted data from ratio analysis
            ratio_data = self.ratio_analysis.prepare_excel_data()
            log_debug(f"\nPrepared {len(ratio_data.get('data', [])) if ratio_data else 0} ratio records")
            
            return ratio_data
            
        except Exception as e:
            log_debug(f"Error preparing ratio data: {str(e)}", include_trace=True)
            return {'headers': [], 'data': []}

    def _get_valid_context_pairs(self) -> List[Tuple[str, str]]:
        """Get valid BS/IS context pairs from data manager with enhanced CF validation"""
        log_debug("\n=== Getting Valid Context Pairs (ExcelExporter._get_valid_context_pairs) ===")
        
        try:
            # Use new context-based matching from data manager
            log_debug("Calling self.data_manager.find_matching_contexts()")
            matching_contexts = self.data_manager.find_matching_contexts()
            if not matching_contexts:
                log_debug("No valid context pairs found (ExcelExporter._get_valid_context_pairs)")
                return []
                
            # Sort pairs by end date
            def get_end_date(ctx_pair):
                bs_ctx, is_ctx = ctx_pair
                bs_period = self.data_manager.fact_provider.filing_analyzer.statement_periods['BS'].get(bs_ctx)
                if bs_period and bs_period[1]:
                    return bs_period[1]
                is_period = self.data_manager.fact_provider.filing_analyzer.statement_periods['IS'].get(is_ctx)
                return is_period[1] if is_period else datetime.min
                
            sorted_pairs = sorted(matching_contexts, key=get_end_date)
            
            # Log sorted pairs with their periods AND CF matches
            log_debug("\nProcessing context pairs with CF:")
            for bs_ctx, is_ctx in sorted_pairs:
                bs_period = self.data_manager.fact_provider.filing_analyzer.statement_periods['BS'].get(bs_ctx)
                is_period = self.data_manager.fact_provider.filing_analyzer.statement_periods['IS'].get(is_ctx)
                
                # Get matching CF contexts
                cf_contexts = self.data_manager.find_matching_cf_contexts(bs_ctx)
                
                log_debug(f"\nContext pair:")
                if bs_period and bs_period[1]:
                    log_debug(f"  BS: {bs_ctx} - End={bs_period[1].strftime('%Y-%m-%d')}")
                if is_period and is_period[1]:
                    log_debug(f"  IS: {is_ctx} - {is_period[0].strftime('%Y-%m-%d')} to {is_period[1].strftime('%Y-%m-%d')}")
                if cf_contexts:
                    log_debug(f"  CF matches: {cf_contexts}")
                else:
                    log_debug("  No matching CF contexts found")
            
            log_debug(f"\nFound {len(sorted_pairs)} valid context pairs")
            return sorted_pairs
            
        except Exception as e:
            log_debug(f"Error getting valid context pairs: {str(e)}", include_trace=True)
            return []

    def _log_export_results(self, successful_exports: List[tuple], failed_exports: List[tuple]):
        """Log export results"""
        log_debug(f"\nSuccessfully exported {len(successful_exports)} sheets:")
        for sheet_name, full_name in successful_exports:
            log_debug(f"- {sheet_name}: {full_name}")

        if failed_exports:
            log_debug("\nFailed exports:")
            for name, reason in failed_exports:
                log_debug(f"- {name}: {reason}")

    def transform_ratio_data_for_columns(self, ratio_data: dict) -> dict:
        """Transform ratio data to period-based column format"""
        log_debug("\n=== Transforming Ratio Data to Column Format ===")
        
        try:
            if not ratio_data or 'data' not in ratio_data:
                log_debug("No ratio data to transform")
                return {'headers': [], 'data': []}
                
            # Extract periods and ratios
            periods = set()
            ratios = {}
            
            log_debug("Processing ratio data rows...")
            for row in ratio_data['data']:
                # if len(row) < 4:  # Basic validation
                #     continue
                if len(row) < 4 or not row[0] or not row[3]:  # Added check for empty ratio_name or stmt_types
                    continue                
                    
                ratio_name = row[0]    # Ratio Name
                value = row[1]         # Value
                period = row[2]        # Period
                stmt_types = row[3]    # Statement Types Used
                
                # Skip rows with empty essential values
                if not ratio_name.strip() or not stmt_types.strip():
                    continue                
                
                periods.add(period)
                if ratio_name not in ratios:
                    ratios[ratio_name] = {
                        'stmt_types': stmt_types,
                        'values': {}
                    }
                ratios[ratio_name]['values'][period] = value
                
            # Create new structure
            new_headers = ['Ratio Name', 'Statements used']
            sorted_periods = sorted(periods)
            new_headers.extend(sorted_periods)
            
            new_data = []
            for ratio_name, info in sorted(ratios.items()):
                row = [ratio_name, info['stmt_types']]
                for period in sorted_periods:
                    row.append(info['values'].get(period, ''))
                new_data.append(row)
                
            log_debug(f"Transformed {len(ratio_data['data'])} rows into {len(new_data)} ratio records")
            log_debug(f"Periods found: {len(periods)}")
            
            return {'headers': new_headers, 'data': new_data}
            
        except Exception as e:
            log_debug(f"Error transforming ratio data: {str(e)}", include_trace=True)
            return {'headers': [], 'data': []}

    
class ValueMetadata:
    """Metadata for tracked financial values"""
    def __init__(self, qname: str = None, source_statement: str = None, 
                 is_custom: bool = False, components: Dict[str, float] = None, 
                 calculation_rule: str = None, last_updated: datetime = None,
                 context_id: str = None):  # Added context_id
        self.qname = qname
        self.source_statement = source_statement
        self.is_custom = is_custom
        self.components = components or {}
        self.calculation_rule = calculation_rule
        self.last_updated = last_updated or datetime.now()
        self.context_id = context_id  # Store contextRef
  
    def to_dict(self) -> dict:
        """Convert metadata to dictionary for logging"""
        return {
            'qname': self.qname,
            'source': self.source_statement,
            'is_custom': self.is_custom,
            'components': {k: {'value': v} for k, v in self.components.items()} if self.components else {},
            'calculation_rule': self.calculation_rule,
            'last_updated': self.last_updated.isoformat() if self.last_updated else None,
            'context_id': self.context_id  # Include contextRef in output
        }
        
        
class PluginTester:
    """Framework for testing ARC Plugin functionality"""
    
    def __init__(self, model_xbrl):
        self.model_xbrl = model_xbrl
        self.data_manager = model_xbrl._fact_provider.financial_data_manager
        self.logger = logger
        self.test_results = defaultdict(list)
        
    def run_all_tests(self):
        """Run all test suites"""
        try:
            self.test_core_functionality()
            self.test_data_flow()
            self.test_output_validation()
            self.report_results()
        except Exception as e:
            self.logger.log(f"Testing error: {str(e)}", include_trace=True)

    def test_core_functionality(self):
        """Test core plugin functionality"""
        self.logger.log("\n=== Testing Core Functionality ===")
        
        # Test statement detection
        self._test_statement_detection()
        
        # Test period handling
        self._test_period_handling()
        
        # Test context mapping
        self._test_context_mapping()

    def _test_statement_detection(self):
        """Test financial statement detection"""
        test_cases = [
            ("Statement of Financial Position", "BS"),
            ("Balance Sheet", "BS"),
            ("Konzernbilanz", "BS"),
            ("Income Statement", "IS"),
            ("Gewinn- und Verlustrechnung", "IS"),
            ("Statement of Cash Flows", "CF")
        ]
        
        detector = StatementTypeDetector()
        for definition, expected in test_cases:
            stmt_type, confidence = detector.detect_statement_type(definition)
            self.test_results["statement_detection"].append({
                "definition": definition,
                "expected": expected,
                "actual": stmt_type,
                "confidence": confidence,
                "passed": stmt_type == expected
            })

    def _test_period_handling(self):
        """Test period handling and matching"""
        matching_periods = self.data_manager.find_matching_periods()
        
        for bs_period, is_period in matching_periods:
            # Verify period format
            self._verify_period(bs_period, "BS")
            self._verify_period(is_period, "IS")
            
            # Verify matching
            valid = self.data_manager.validate_matching_periods(bs_period, is_period)
            self.test_results["period_matching"].append({
                "bs_period": bs_period,
                "is_period": is_period,
                "valid": valid,
                "passed": valid
            })

    def _verify_period(self, period, stmt_type):
        """Verify period structure"""
        if stmt_type == "BS":
            valid = period[0] is None and period[1] is not None
        else:
            valid = period[0] is not None and period[1] is not None
            
        self.test_results["period_format"].append({
            "period": period,
            "type": stmt_type,
            "valid": valid,
            "passed": valid
        })

    def test_data_flow(self):
        """Test data flow and value tracking"""
        self.logger.log("\n=== Testing Data Flow ===")
        
        # Test value retrieval
        self._test_value_retrieval()
        
        # Test compound values
        self._test_compound_values()
        
        # Test caching
        self._test_caching()

    def _test_value_retrieval(self):
        """Test basic value retrieval"""
        test_concepts = ['total_assets', 'revenue', 'net_income']
        periods = self.data_manager.find_matching_periods()
        
        for concept in test_concepts:
            for bs_period, is_period in periods:
                period = bs_period if concept == 'total_assets' else is_period
                stmt_type = 'BS' if concept == 'total_assets' else 'IS'
                
                value, metadata = self.data_manager.get_value(concept, period, stmt_type)
                self.test_results["value_retrieval"].append({
                    "concept": concept,
                    "period": period,
                    "value": value,
                    "has_metadata": metadata is not None,
                    "passed": value is not None and metadata is not None
                })

    def _test_compound_values(self):
        """Test compound value calculation"""
        test_cases = [
            ("total_debt", "BS"),
            ("total_capital", "BS"),
            ("gross_profit", "IS")
        ]
        
        periods = self.data_manager.find_matching_periods()
        for concept, stmt_type in test_cases:
            for bs_period, is_period in periods:
                period = bs_period if stmt_type == "BS" else is_period
                value, metadata = self.data_manager.get_value(concept, period, stmt_type)
                
                composition = self.data_manager.get_value_composition(concept, period)
                self.test_results["compound_values"].append({
                    "concept": concept,
                    "period": period,
                    "has_value": value is not None,
                    "has_composition": composition is not None,
                    "passed": value is not None and composition is not None
                })

    def test_output_validation(self):
        """Test Excel output format and data consistency"""
        self.logger.log("\n=== Testing Output Validation ===")
        
        # Get source data
        excel_data = self.model_xbrl._exporter.prepare_excel_data()
        
        # Validate statements
        self._validate_statements(excel_data.get('statements', []))
        
        # Validate ratios
        self._validate_ratios(excel_data.get('ratios', {}))
        
        # Validate inputs
        self._validate_inputs(excel_data.get('inputs', {}))

    def _validate_statements(self, statements):
        """Validate statement data"""
        for stmt in statements:
            self.test_results["statement_validation"].append({
                "name": stmt['statement_name'],
                "has_headers": len(stmt['headers']) > 0,
                "has_data": len(stmt['data']) > 0,
                "passed": len(stmt['headers']) > 0 and len(stmt['data']) > 0
            })

    def _validate_ratios(self, ratio_data):
        """Validate ratio calculations"""
        if not ratio_data:
            return
            
        headers = ratio_data.get('headers', [])
        data = ratio_data.get('data', [])
        
        self.test_results["ratio_validation"].append({
            "has_headers": len(headers) > 0,
            "has_data": len(data) > 0,
            "consistent_columns": all(len(row) == len(headers) for row in data),
            "passed": len(headers) > 0 and len(data) > 0
        })

    def report_results(self):
        """Generate test report"""
        self.logger.log("\n=== Test Results ===")
        
        total_tests = 0
        passed_tests = 0
        
        for category, results in self.test_results.items():
            category_total = len(results)
            category_passed = sum(1 for r in results if r["passed"])
            
            self.logger.log(f"\n{category}:")
            self.logger.log(f"Passed: {category_passed}/{category_total}")
            
            for result in results:
                if not result["passed"]:
                    self.logger.log(f"Failed test: {result}")
                    
            total_tests += category_total
            passed_tests += category_passed
            
        self.logger.log(f"\nOverall Results:")
        self.logger.log(f"Total Tests: {total_tests}")
        self.logger.log(f"Passed: {passed_tests}")
        self.logger.log(f"Success Rate: {(passed_tests/total_tests)*100:.2f}%")




class ConfigHandler:
    """Manages plugin configuration and settings. No GUI handling is done here"""
    
    DEFAULT_CONFIG = {
        "excel_export": {
            "local_files": {
                "mode": "central_folder",   # "off", "filing_folder", "central_folder"
                "central_folder": r"C:\Data\excel_output"
            },
            "online_files": {
                "mode": "central_folder", # "off", "central_folder" 
                "central_folder": r"C:\Data\online_output"
            }
        },
        "logging": {
            "enabled": True
        }
    }
    
    def __init__(self, plugin_dir):
        self.plugin_dir = plugin_dir
        self.config_path = os.path.join(plugin_dir, 'config.json')
        self.config = self._load_config()
        
    def _load_config(self):
        """Load config from file or create default"""
        try:
            if os.path.exists(self.config_path):
                with open(self.config_path, 'r') as f:
                    return json.load(f)
            return self._create_default_config()
        except Exception as e:
            log_debug(f"Error loading config: {str(e)}")
            return self._create_default_config()
            
    def _create_default_config(self):
        """Create and save default config"""
        try:
            os.makedirs(self.plugin_dir, exist_ok=True)
            with open(self.config_path, 'w') as f:
                json.dump(self.DEFAULT_CONFIG, f, indent=4)
            return self.DEFAULT_CONFIG.copy()
        except Exception as e:
            log_debug(f"Error creating config: {str(e)}")
            return self.DEFAULT_CONFIG.copy()
            
    def save_config(self):
        """Save current config to file"""
        try:
            with open(self.config_path, 'w') as f:
                json.dump(self.config, f, indent=4)
        except Exception as e:
            log_debug(f"Error saving config: {str(e)}")
            
    def get_export_path(self, model_xbrl):
        """Get export path based on config and filing type"""
        try:
            is_online = model_xbrl.uri.startswith(('http://', 'https://'))
            config_section = 'online_files' if is_online else 'local_files'
            mode = self.config['excel_export'][config_section]['mode']
            
            if mode == 'off':
                return None
                
            if not is_online and mode == 'filing_folder':
                return os.path.dirname(model_xbrl.uri)
                
            return self.config['excel_export'][config_section]['central_folder']
            
        except Exception as e:
            log_debug(f"Error getting export path: {str(e)}")
            return None
            
    def is_logging_enabled(self):
        """Check if logging is enabled"""
        enabled = self.config.get('logging', {}).get('enabled', True)
        print(f"\nChecking logging config:")
        print(f"Full config: {self.config}")
        print(f"Logging enabled: {enabled}")
        return enabled


class ConfigDialog(tk.Toplevel):
    """Dialog for editing plugin configuration"""
    
    def __init__(self, parent, config_handler):
        super().__init__(parent)
        self.config_handler = config_handler
        self.result = None
        
        self.title("ARC Plugin Configuration")
        self.geometry("600x400")
        
        # Create main frame with padding
        main_frame = ttk.Frame(self, padding="10")
        main_frame.grid(row=0, column=0, sticky=(tk.W, tk.E, tk.N, tk.S))
        
        # Local Files Frame
        local_frame = ttk.LabelFrame(main_frame, text="Excel Export for Local Filings (ZIP)", padding="5")
        local_frame.grid(row=0, column=0, sticky=(tk.W, tk.E), pady=5)
        
        # Local Mode
        ttk.Label(local_frame, text="Export Mode:").grid(row=0, column=0, sticky=tk.W)
        self.local_mode = tk.StringVar(value=self.config_handler.config['excel_export']['local_files']['mode'])
        local_mode_cb = ttk.Combobox(local_frame, textvariable=self.local_mode, 
                                    values=["off", "filing_folder", "central_folder"],
                                    state="readonly", width=20)
        local_mode_cb.grid(row=0, column=1, sticky=tk.W, padx=5)
        
        # Local Folder
        ttk.Label(local_frame, text="Central Folder:").grid(row=1, column=0, sticky=tk.W)
        self.local_folder = tk.StringVar(value=self.config_handler.config['excel_export']['local_files']['central_folder'])
        local_folder_entry = ttk.Entry(local_frame, textvariable=self.local_folder, width=50)
        local_folder_entry.grid(row=1, column=1, sticky=tk.W, padx=5)
        ttk.Button(local_frame, text="Browse...", 
                   command=lambda: self.browse_folder(self.local_folder)).grid(row=1, column=2)
        
        # Online Files Frame
        online_frame = ttk.LabelFrame(main_frame, text="Excel Export for Online Filings (URL)", padding="5")
        online_frame.grid(row=1, column=0, sticky=(tk.W, tk.E), pady=5)
        
        # Online Mode
        ttk.Label(online_frame, text="Export Mode:").grid(row=0, column=0, sticky=tk.W)
        self.online_mode = tk.StringVar(value=self.config_handler.config['excel_export']['online_files']['mode'])
        online_mode_cb = ttk.Combobox(online_frame, textvariable=self.online_mode,
                                     values=["off", "central_folder"],
                                     state="readonly", width=20)
        online_mode_cb.grid(row=0, column=1, sticky=tk.W, padx=5)
        
        # Online Folder
        ttk.Label(online_frame, text="Central Folder:").grid(row=1, column=0, sticky=tk.W)
        self.online_folder = tk.StringVar(value=self.config_handler.config['excel_export']['online_files']['central_folder'])
        online_folder_entry = ttk.Entry(online_frame, textvariable=self.online_folder, width=50)
        online_folder_entry.grid(row=1, column=1, sticky=tk.W, padx=5)
        ttk.Button(online_frame, text="Browse...", 
                   command=lambda: self.browse_folder(self.online_folder)).grid(row=1, column=2)
        
        # Logging Frame
        logging_frame = ttk.LabelFrame(main_frame, text="Write Log Files", padding="5")
        logging_frame.grid(row=2, column=0, sticky=(tk.W, tk.E), pady=5)
        
        # Logging Enabled
        self.logging_enabled = tk.BooleanVar(value=self.config_handler.config['logging']['enabled'])
        ttk.Checkbutton(logging_frame, text="Enable Logging", 
                       variable=self.logging_enabled).grid(row=0, column=0, sticky=tk.W)
        
        # Buttons
        button_frame = ttk.Frame(main_frame)
        button_frame.grid(row=3, column=0, sticky=(tk.E), pady=10)
        ttk.Button(button_frame, text="Save", command=self.save).grid(row=0, column=0, padx=5)
        ttk.Button(button_frame, text="Cancel", command=self.cancel).grid(row=0, column=1)
        
        # Update states based on initial values
        self.update_states()
        
        # Bind mode changes to state updates
        local_mode_cb.bind('<<ComboboxSelected>>', lambda e: self.update_states())
        online_mode_cb.bind('<<ComboboxSelected>>', lambda e: self.update_states())
        
        # Make dialog modal
        self.transient(parent)
        self.grab_set()
        
    def update_states(self):
        """Update entry states based on mode selections"""
        # Local folder state
        local_central = self.local_mode.get() == "central_folder"
        for child in self.winfo_children():
            if isinstance(child, ttk.Entry) and child.winfo_parent().endswith('local_frame'):
                child['state'] = 'normal' if local_central else 'disabled'
                
        # Online folder state
        online_central = self.online_mode.get() == "central_folder"
        for child in self.winfo_children():
            if isinstance(child, ttk.Entry) and child.winfo_parent().endswith('online_frame'):
                child['state'] = 'normal' if online_central else 'disabled'
    
    def browse_folder(self, var):
        """Open folder browser and update entry"""
        folder = filedialog.askdirectory(initialdir=var.get())
        if folder:
            var.set(folder)
            
    def save(self):
        """Save configuration changes"""
        try:
            config = self.config_handler.config
            
            # Update local files config
            config['excel_export']['local_files']['mode'] = self.local_mode.get()
            config['excel_export']['local_files']['central_folder'] = self.local_folder.get()
            
            # Update online files config
            config['excel_export']['online_files']['mode'] = self.online_mode.get()
            config['excel_export']['online_files']['central_folder'] = self.online_folder.get()
            
            # Update logging config
            config['logging']['enabled'] = self.logging_enabled.get()
            
            # Save to file
            self.config_handler.save_config()
            
            self.result = True
            self.destroy()
            
        except Exception as e:
            messagebox.showerror("Error", f"Failed to save configuration: {str(e)}")
            
    def cancel(self):
        """Cancel without saving"""
        self.result = False
        self.destroy()




# Helper functions

def get_statement_patterns():
    """Returns dictionary of patterns for financial statements with multilingual support"""
    return {
        'bs': {  # Balance Sheet patterns
            'primary': [  
                r".*?(statement.*?financial.*?position|balance.*?sheet|bilanz|vermögenslage|finanzlage)",
                r".*?(consolidated.*?statement.*?financial.*?position)",
                r".*?(assets.*?liabilities|aktiva.*?passiva)"
            ],
            'en': [
                r".*?statement.*?financial.*?position",
                r".*?balance.*?sheet",
                r".*?statement.*?(assets|liabilities)",
                r".*?consolidated.*?balance.*?sheet",
                r".*?consolidated.*?statement.*?financial.*?position"
            ],
            'de': [
                r".*?bilanz",
                r".*?vermögenslage",
                r".*?finanzlage",
                r".*?jahresbilanz",
                r".*?konzernbilanz",
                r".*?konzern.*?bilanz"
            ],
            'mixed': [  
                r".*?(balance.*?sheet|bilanz)",
                r".*?(statement.*?financial.*?position|vermögenslage)",
                r".*?(assets.*?liabilities|aktiva.*?passiva)"
            ],
            'labels': ["BS", "Statement of Financial Position", "Bilanz"]
        },
        'pl': {  # Profit & Loss patterns
            'primary': [
                r".*?(statement.*?(profit|loss|income)|income.*?statement|gewinn.*?verlust|ergebnisrechnung)"
            ],
            'en': [
                r".*?statement.*?(profit|loss)",
                r".*?income.*?statement",
                r".*?profit.*?loss.*?account",
                r".*?statement.*?operations",
                r".*?consolidated.*?income.*?statement",
                r".*?consolidated.*?statement.*?(profit|loss)"
            ],
            'de': [
                r".*?gewinn.*?verlust.*?(rechnung|konto)",
                r".*?ergebnisrechnung",
                r".*?erfolgsrechnung",
                r".*?jahresüberschuss",
                r".*?konzern.*?gewinn.*?verlust.*?rechnung",
                r".*?guv"
            ],
            'mixed': [
                r".*?(income.*?statement|gewinn.*?verlust)",
                r".*?(profit.*?loss|ergebnisrechnung)"
            ],
            'labels': ["IS", "Income Statement", "GuV"]
        },
        'ci': {  # Comprehensive Income patterns
            'primary': [
                r".*?(comprehensive.*?income|gesamtergebnis|gesamtergebnisrechnung)"
            ],
            'en': [
                r".*?comprehensive.*?income",
                r".*?statement.*?comprehensive.*?income",
                r".*?total.*?comprehensive.*?income",
                r".*?consolidated.*?comprehensive.*?income"
            ],
            'de': [
                r".*?gesamtergebnis",
                r".*?gesamtergebnisrechnung",
                r".*?gesamtperiodenerfolg",
                r".*?konzern.*?gesamtergebnis",
                r".*?konzern.*?gesamtergebnisrechnung"
            ],
            'mixed': [
                r".*?(comprehensive.*?income|gesamtergebnis)",
                r".*?(total.*?comprehensive.*?income|gesamtergebnisrechnung)"
            ],
            'abbreviated': [  # Last resort patterns
                r".*?\bCI\b",
                r".*?consolidated\s+CI\b",
                r".*?konzern\s+CI\b",
                r".*?CI[\s-]*tabelle\b"
            ],
            'labels': ["CI", "Statement of Comprehensive Income", "Gesamtergebnisrechnung"]
        },
        'cf': {  # Cash Flow patterns
            'primary': [
                r".*?(cash.*?flow|kapitalfluss|cashflow|geldfluss)"
            ],
            'en': [
                r".*?cash.*?flow.*?statement",
                r".*?statement.*?cash.*?flows?",
                r".*?cash.*?flow.*?analysis",
                r".*?consolidated.*?cash.*?flows?",
                r".*?consolidated.*?statement.*?cash.*?flows?"
            ],
            'de': [
                r".*?kapitalflussrechnung",
                r".*?cashflow.*?rechnung",
                r".*?geldflussrechnung",
                r".*?finanzmittelflussrechnung",
                r".*?konzern.*?kapitalflussrechnung"
            ],
            'mixed': [
                r".*?(cash.*?flow|kapitalfluss)",
                r".*?(statement.*?cash.*?flows?|kapitalflussrechnung)"
            ],
            'abbreviated': [  # Last resort patterns
                r".*?\bCF\b",
                r".*?consolidated\s+CF\b",
                r".*?konzern\s+CF\b",
                r".*?CF[\s-]*tabelle\b"
            ],
            'labels': ["CF", "Cash Flow Statement", "Kapitalflussrechnung"]
        },
        'eq': {  # Equity patterns
            'primary': [
                r".*?(changes.*?equity|eigenkapitalveränderung|eigenkapital.*?entwicklung)"
            ],
            'en': [
                r".*?changes.*?equity",
                r".*?statement.*?changes.*?equity",
                r".*?statement.*?stockholders.*?equity",
                r".*?consolidated.*?changes.*?equity",
                r".*?consolidated.*?statement.*?changes.*?equity"
            ],
            'de': [
                r".*?eigenkapitalveränderungsrechnung",
                r".*?eigenkapital.*?entwicklung",
                r".*?entwicklung.*?eigenkapital",
                r".*?konzern.*?eigenkapitalveränderungsrechnung",
                r".*?konzern.*?eigenkapital.*?entwicklung"
            ],
            'mixed': [
                r".*?(changes.*?equity|eigenkapitalveränderung)",
                r".*?(statement.*?changes.*?equity|eigenkapital.*?entwicklung)"
            ],
            'abbreviated': [  # Last resort patterns
                r".*?\bCE\b",
                r".*?consolidated\s+CE\b",
                r".*?konzern\s+CE\b",
                r".*?CE[\s-]*tabelle\b"
            ],
            'labels': ["EQ", "Statement of Changes in Equity", "Eigenkapitalveränderungsrechnung"]
        },
        'notes': {  # Notes patterns
            'primary': [
                r".*?(notes|anhang|erläuterungen|disclosures)"
            ],
            'en': [
                r".*?notes.*?financial.*?statements?",
                r".*?disclosures",
                r".*?explanatory.*?notes",
                r".*?consolidated.*?notes"
            ],
            'de': [
                r".*?anhang",
                r".*?erläuterungen",
                r".*?konzernanhang",
                r".*?notes"  # German companies often use English "notes"
            ],
            'mixed': [
                r".*?(notes|anhang)",
                r".*?(disclosures|erläuterungen)"
            ],
            'labels': ["NOT", "Notes", "Anhang"]
        },
        'accounting_policies': {  # Accounting Policies patterns
            'primary': [
                r".*?(accounting.*?policies|bilanzierungs.*?bewertung)"
            ],
            'en': [
                r".*?accounting.*?policies",
                r".*?significant.*?accounting.*?policies",
                r".*?accounting.*?principles",
                r".*?significant.*?accounting.*?principles"
            ],
            'de': [
                r".*?bilanzierungs.*?bewertungsmethoden",
                r".*?rechnungslegungsmethoden",
                r".*?bilanzierungs.*?bewertungsgrundsätze",
                r".*?grundsätze.*?bilanzierung"
            ],
            'mixed': [
                r".*?(accounting.*?policies|bilanzierungsmethoden)",
                r".*?(accounting.*?principles|bewertungsgrundsätze)"
            ],
            'labels': ["POL", "Accounting Policies", "Bilanzierungs- und Bewertungsmethoden"]
        }
    }

def get_statement_type_order():
    """
    Central definition of statement type ordering for both IFRS and German GAAP filings.
    Returns a dictionary with statement type codes and their display order.
    
    Ordering follows common financial reporting practices:
    1. Balance Sheet (BS) first
    2. Income Statement (IS) / Profit & Loss
    3. Comprehensive Income (CI)
    4. Changes in Equity (EQ)
    5. Cash Flow (CF)
    
    Each main statement type can have a parenthetical version (P suffix)
    which shares the same base order number but is sorted after the main version.
    """
    return {
        # Balance Sheet group - order 0
        'BS': 0,    # Balance Sheet / Statement of Financial Position
        'BSP': 0,   # Balance Sheet (parenthetical)
        
        # Income Statement group - order 1
        'PL': 1,    # Profit/Loss (alternative name)
        'IS': 1,    # Income Statement
        'ISP': 1,   # Income Statement (parenthetical)
        
        # Comprehensive Income group - order 2
        'CI': 2,    # Statement of Comprehensive Income
        'CIP': 2,   # Comprehensive Income (parenthetical)
        'SCI': 2,   # Statement of Comprehensive Income (alternative name)
        'SCIP': 2,  # Statement of Comprehensive Income (parenthetical)
        
        # Equity group - order 3
        'EQ': 3,    # Statement of Changes in Equity
        'EQP': 3,   # Equity Statement (parenthetical)
        'SCE': 3,   # Statement of Changes in Equity (alternative name)
        'SCEP': 3,  # Statement of Changes in Equity (parenthetical)
        
        # Cash Flow group - order 4
        'CF': 4,    # Cash Flow Statement
        'CFP': 4,   # Cash Flow Statement (parenthetical)
        'SCF': 4,   # Statement of Cash Flows (alternative name)
        'SCFP': 4,  # Statement of Cash Flows (parenthetical)
        
        # Other statements - order 5
        'OTH': 5,   # Other Financial Statements
        'OTHP': 5,  # Other Financial Statements (parenthetical)
        
        # Special cases - higher orders
        'DEI': 6,   # Document and Entity Information
        'LAB': 7,   # Labels
        'CAL': 8,   # Calculations
        'PRE': 9,   # Presentations
        'DEF': 10   # Definitions
    }

def sort_statements(statements: List[Tuple[str, object]], is_usgaap: bool = False) -> List[Tuple[str, object]]:
    """
    Sort statements based on type, handling both IFRS and US-GAAP styles.
    
    Args:
        statements: List of tuples containing (definition, roleType)
        is_usgaap: Boolean indicating if this is a US-GAAP filing
        
    Returns:
        Sorted list of statement tuples
        
    Sort criteria:
    1. Primary order based on statement type (BS -> IS -> CI -> EQ -> CF)
    2. Within same type, non-parenthetical before parenthetical
    3. Within same group, original definition order as tie-breaker
    4. Only statements marked as financial statements are included
    """
    log_debug("\nSorting Financial Statements:")
    
    # Get the order mapping
    statement_type_order = get_statement_type_order()
    
    def get_sort_key(stmt_tuple):
        definition, role_type = stmt_tuple
        
        # Log the statement being processed
        log_debug(f"\nProcessing statement for sorting:")
        log_debug(f"Definition: {definition}")
        
        # Get type based on filing style
        if is_usgaap:
            stmt_type = getattr(role_type, '_tableCode', None)
            log_debug(f"US-GAAP statement type from _tableCode: {stmt_type}")
        else:
            stmt_type = getattr(role_type, '_statementType', None)
            log_debug(f"IFRS statement type from _statementType: {stmt_type}")
            
        # Get the table index group (to ensure it's a financial statement)
        table_index = getattr(role_type, '_tableIndex', None)
        group = table_index[0] if table_index else None
        
        # Skip if not a financial statement
        if group != "2Financial Statements":
            log_debug("Skipping - not a financial statement")
            return (999, True, definition)
        
        # Get order number from mapping
        order = statement_type_order.get(stmt_type, 999) if stmt_type else 999
        log_debug(f"Order number: {order}")
        
        # Check if parenthetical
        is_parenthetical = stmt_type and stmt_type.endswith('P') if stmt_type else False
        log_debug(f"Is parenthetical: {is_parenthetical}")
        
        # Get sequence number if available (for tie-breaking)
        seq_num = table_index[1] if table_index else "999"
        log_debug(f"Sequence number: {seq_num}")
        
        # Return sort key tuple
        return (order, is_parenthetical, seq_num, definition)
    
    # Filter for financial statements and sort
    filtered_statements = [
        stmt for stmt in statements 
        if hasattr(stmt[1], '_tableIndex') and 
           stmt[1]._tableIndex and 
           stmt[1]._tableIndex[0] == "2Financial Statements"
    ]
    
    log_debug(f"\nFound {len(filtered_statements)} financial statements to sort")
    
    # Sort statements
    sorted_statements = sorted(filtered_statements, key=get_sort_key)
    
    # Log the final order
    log_debug("\nFinal statement order:")
    for definition, role_type in sorted_statements:
        stmt_type = (getattr(role_type, '_tableCode', None) if is_usgaap 
                    else getattr(role_type, '_statementType', None))
        log_debug(f"- {definition} (Type: {stmt_type})")
    
    return sorted_statements

def test_statement_detection(definition: str) -> str:
    """
    Test function to diagnose statement type detection with detailed logging
    
    Args:
        definition: Role definition string to analyze
        
    Returns:
        Detected statement type code
    
    The function provides detailed logging of:
    1. Pattern matching results
    2. Confidence scores
    3. Language detection
    4. Pattern category matches
    5. Special case handling
    """
    log_debug(f"\nTesting statement detection for: '{definition}'")
    log_debug("=" * 80)
    
    definition_lower = definition.lower()
    
    # Get all statement patterns
    statement_patterns = get_statement_patterns()
    
    # Track matches and scores for each statement type
    matches = defaultdict(lambda: {
        'primary': [],
        'en': [],
        'de': [],
        'mixed': [],
        'abbreviated': [],
        'indicators': [],
        'score': 0.0
    })
    
    # Check each statement type's patterns
    for stmt_type, patterns in statement_patterns.items():
        log_debug(f"\nChecking {stmt_type.upper()} patterns:")
        
        # Check primary patterns
        if 'primary' in patterns:
            log_debug("\nPrimary patterns:")
            for pattern in patterns['primary']:
                match = bool(re.match(pattern, definition_lower, re.I))
                log_debug(f"  {pattern}: {'Match' if match else 'No match'}")
                if match:
                    matches[stmt_type]['primary'].append(pattern)
                    matches[stmt_type]['score'] += 0.4
        
        # Check English patterns
        if 'en' in patterns:
            log_debug("\nEnglish patterns:")
            for pattern in patterns['en']:
                match = bool(re.match(pattern, definition_lower, re.I))
                log_debug(f"  {pattern}: {'Match' if match else 'No match'}")
                if match:
                    matches[stmt_type]['en'].append(pattern)
                    matches[stmt_type]['score'] += 0.3
                    
        # Check German patterns
        if 'de' in patterns:
            log_debug("\nGerman patterns:")
            for pattern in patterns['de']:
                match = bool(re.match(pattern, definition_lower, re.I))
                log_debug(f"  {pattern}: {'Match' if match else 'No match'}")
                if match:
                    matches[stmt_type]['de'].append(pattern)
                    matches[stmt_type]['score'] += 0.3
                    
        # Check mixed language patterns
        if 'mixed' in patterns:
            log_debug("\nMixed language patterns:")
            for pattern in patterns['mixed']:
                match = bool(re.match(pattern, definition_lower, re.I))
                log_debug(f"  {pattern}: {'Match' if match else 'No match'}")
                if match:
                    matches[stmt_type]['mixed'].append(pattern)
                    matches[stmt_type]['score'] += 0.2
                    
        # Check abbreviated patterns
        if 'abbreviated' in patterns:
            log_debug("\nAbbreviated patterns:")
            for pattern in patterns['abbreviated']:
                match = bool(re.match(pattern, definition_lower, re.I))
                log_debug(f"  {pattern}: {'Match' if match else 'No match'}")
                if match:
                    matches[stmt_type]['abbreviated'].append(pattern)
                    matches[stmt_type]['score'] += 0.1
                    
        # Check indicators
        if 'indicators' in patterns:
            log_debug("\nIndicators:")
            for indicator in patterns['indicators']:
                if indicator.lower() in definition_lower:
                    log_debug(f"  Found indicator: {indicator}")
                    matches[stmt_type]['indicators'].append(indicator)
                    matches[stmt_type]['score'] += 0.1
                    
    # Check for parenthetical
    is_parenthetical = bool(re.search(r"pa?r[ae]ne?th\w?[aei]+\w?t?h?i?c", definition_lower))
    log_debug(f"\nParenthetical check: {'Yes' if is_parenthetical else 'No'}")
    
    # Log match summary
    log_debug("\nMatch Summary:")
    log_debug("-" * 40)
    for stmt_type, match_info in matches.items():
        if match_info['score'] > 0:
            log_debug(f"\n{stmt_type.upper()} (Score: {match_info['score']:.2f}):")
            if match_info['primary']:
                log_debug("  Primary matches:")
                for pattern in match_info['primary']:
                    log_debug(f"    - {pattern}")
            if match_info['en']:
                log_debug("  English matches:")
                for pattern in match_info['en']:
                    log_debug(f"    - {pattern}")
            if match_info['de']:
                log_debug("  German matches:")
                for pattern in match_info['de']:
                    log_debug(f"    - {pattern}")
            if match_info['mixed']:
                log_debug("  Mixed matches:")
                for pattern in match_info['mixed']:
                    log_debug(f"    - {pattern}")
            if match_info['abbreviated']:
                log_debug("  Abbreviated matches:")
                for pattern in match_info['abbreviated']:
                    log_debug(f"    - {pattern}")
            if match_info['indicators']:
                log_debug("  Indicators found:")
                for indicator in match_info['indicators']:
                    log_debug(f"    - {indicator}")
    
    # Get best match
    best_match = max(matches.items(), key=lambda x: x[1]['score'], default=(None, {'score': 0}))
    best_type = best_match[0]
    best_score = best_match[1]['score']
    
    # Add parenthetical suffix if needed
    if is_parenthetical and best_type:
        best_type += 'P'
    
    # Log final determination
    log_debug("\nFinal Determination:")
    log_debug("-" * 40)
    log_debug(f"Best matching type: {best_type or 'None'}")
    log_debug(f"Confidence score: {best_score:.2f}")
    log_debug(f"Parenthetical: {'Yes' if is_parenthetical else 'No'}")
    
    return best_type if best_score >= 0.3 else 'OTH'

def debug_statement_detection(model_xbrl):
    """
    Run detection tests on all role definitions in the model
    """
    log_debug("\nTesting statement detection on all definitions:")
    log_debug("=" * 80)
    
    found_statements = []
    
    # Get all role definitions
    relationship_set = model_xbrl.relationshipSet(XbrlConst.parentChild)
    for roleURI in relationship_set.linkRoleUris:
        for roleType in model_xbrl.roleTypes.get(roleURI, ()):
            definition = roleType.definition
            if definition:
                stmt_type = test_statement_detection(definition)
                if stmt_type != 'OTH':
                    found_statements.append((definition, stmt_type))
    
    # Log summary
    log_debug("\nStatement Detection Summary:")
    log_debug("-" * 40)
    log_debug(f"Found {len(found_statements)} statements:")
    for definition, stmt_type in sorted(found_statements, 
                                      key=lambda x: get_statement_type_order().get(x[1].rstrip('P'), 999)):
        log_debug(f"- {definition} - Type: {stmt_type}")
    
    return found_statements

def test_sort_statements(model_xbrl):
    """
    Test function to verify statement sorting logic and order
    
    Args:
        model_xbrl: The ModelXbrl instance to test
        
    Returns:
        Tuple containing original, IFRS-sorted, and US-GAAP-sorted statements
    """
    log_debug("\n=== Testing Statement Sorting ===")
    
    # Collect all statements
    statements = []
    sorting_info = []  # For tracking sorting criteria
    
    relationship_set = model_xbrl.relationshipSet(XbrlConst.parentChild)
    for roleURI in relationship_set.linkRoleUris:
        for roleType in model_xbrl.roleTypes.get(roleURI, ()):
            if hasattr(roleType, '_tableIndex'):
                statements.append((roleType.definition, roleType))
                
                # Collect sorting criteria for debugging
                sorting_info.append({
                    'definition': roleType.definition,
                    'table_code': getattr(roleType, '_tableCode', 'None'),
                    'statement_type': getattr(roleType, '_statementType', 'None'),
                    'table_index': getattr(roleType, '_tableIndex', ('', '', '')),
                    'is_financial_stmt': getattr(roleType, '_tableIndex', ('',))[0] == "2Financial Statements"
                })
    
    log_debug(f"\nFound {len(statements)} total statements to test")
    
    # Log initial state
    log_debug("\nInitial Statement Information:")
    for info in sorting_info:
        log_debug(f"\nDefinition: {info['definition']}")
        log_debug(f"  Table Code: {info['table_code']}")
        log_debug(f"  Statement Type: {info['statement_type']}")
        log_debug(f"  Table Index: {info['table_index']}")
        log_debug(f"  Is Financial Statement: {info['is_financial_stmt']}")
    
    # Test IFRS sorting
    log_debug("\nTesting IFRS-style sorting:")
    ifrs_sorted = sort_statements(statements, is_usgaap=False)
    log_debug("\nIFRS Sort Result:")
    for i, (definition, role_type) in enumerate(ifrs_sorted, 1):
        stmt_type = getattr(role_type, '_statementType', None)
        order_num = get_statement_type_order().get(stmt_type, 999)
        log_debug(f"{i}. {definition}")
        log_debug(f"   Type: {stmt_type}")
        log_debug(f"   Order: {order_num}")
    
    # Test US-GAAP sorting
    log_debug("\nTesting US-GAAP-style sorting:")
    usgaap_sorted = sort_statements(statements, is_usgaap=True)
    log_debug("\nUS-GAAP Sort Result:")
    for i, (definition, role_type) in enumerate(usgaap_sorted, 1):
        table_code = getattr(role_type, '_tableCode', None)
        order_num = get_statement_type_order().get(table_code, 999)
        log_debug(f"{i}. {definition}")
        log_debug(f"   Code: {table_code}")
        log_debug(f"   Order: {order_num}")
    
    # Verify sorting logic
    log_debug("\nVerifying Sort Results:")
    
    # Check IFRS sorting
    log_debug("\nIFRS Sort Verification:")
    last_order = -1
    for definition, role_type in ifrs_sorted:
        stmt_type = getattr(role_type, '_statementType', None)
        current_order = get_statement_type_order().get(stmt_type, 999)
        if current_order < last_order:
            log_debug(f"WARNING: Potential IFRS sort issue: {definition}")
            log_debug(f"  Current order {current_order} < last order {last_order}")
        last_order = current_order
    
    # Check US-GAAP sorting
    log_debug("\nUS-GAAP Sort Verification:")
    last_order = -1
    for definition, role_type in usgaap_sorted:
        table_code = getattr(role_type, '_tableCode', None)
        current_order = get_statement_type_order().get(table_code, 999)
        if current_order < last_order:
            log_debug(f"WARNING: Potential US-GAAP sort issue: {definition}")
            log_debug(f"  Current order {current_order} < last order {last_order}")
        last_order = current_order
    
    # Count financial statements in each sort
    ifrs_fin_stmt_count = sum(1 for _, rt in ifrs_sorted 
                             if rt._tableIndex[0] == "2Financial Statements")
    usgaap_fin_stmt_count = sum(1 for _, rt in usgaap_sorted 
                               if rt._tableIndex[0] == "2Financial Statements")
    
    log_debug("\nSort Result Summary:")
    log_debug(f"Original statements: {len(statements)}")
    log_debug(f"IFRS financial statements: {ifrs_fin_stmt_count}")
    log_debug(f"US-GAAP financial statements: {usgaap_fin_stmt_count}")
    
    return statements, ifrs_sorted, usgaap_sorted



    # working

def enhance_table_structure(model_xbrl):
    """Enhanced two-phase table structure setup with proper view properties"""
    try:
        log_debug("\n=== Starting Enhanced Table Structure Setup ===")
        
        # Initialize analyzer
        analyzer = FilingAnalyzer(model_xbrl)
        first_table_linkrole = None
        financial_statements = []

        # Phase 1: Complete Role Type Analysis
        log_debug("\nPhase 1: Role Type Analysis")
        relationship_set = model_xbrl.relationshipSet(XbrlConst.parentChild)
        
        for roleURI in relationship_set.linkRoleUris:
            role_types = model_xbrl.roleTypes.get(roleURI, [])
            if not role_types:
                continue
                
            role_type = role_types[0]
            if not role_type.definition:
                continue
                
            # Analyze role type
            analyzer.analyze_role_type(role_type.definition, role_type)
            
            # Track financial statements
            if (hasattr(role_type, '_tableIndex') and 
                role_type._tableIndex[0] == "2Financial Statements"):
                financial_statements.append(role_type)
                if not first_table_linkrole:
                    first_table_linkrole = role_type.roleURI

        # Phase 2: Complete View Setup
        log_debug("\nPhase 2: View Setup")
        model_xbrl._tableFacts = {}
        model_xbrl._tableCodes = {}
        model_xbrl._tableContexts = {}  # Add context tracking
        
        # Set up view properties for each financial statement
        for role_type in financial_statements:
            # Ensure all required properties are set
            if not hasattr(role_type, '_viewProperties'):
                role_type._viewProperties = {}
                
            # Set comprehensive view properties
            role_type._viewProperties.update({
                'label': role_type.definition,
                'roleURI': role_type.roleURI,
                'tableCode': getattr(role_type, '_tableCode', ''),
                'tableIndex': getattr(role_type, '_tableIndex', ('', '', ''))[1],
                'tableFacts': getattr(role_type, '_tableFacts', set()),
                'contexts': getattr(role_type, '_tableContexts', {}),
                'periods': list(getattr(role_type, '_tableContexts', {}).values())
            })
            
            # Store facts and codes at model level
            model_xbrl._tableFacts[role_type.roleURI] = role_type._tableFacts
            model_xbrl._tableCodes[role_type.roleURI] = role_type._tableCode
            model_xbrl._tableContexts[role_type.roleURI] = role_type._tableContexts
            
            log_debug(f"Set up view for: {role_type.definition}")
            log_debug(f"- Code: {role_type._tableCode}")
            log_debug(f"- Facts: {len(role_type._tableFacts)}")
            log_debug(f"- Contexts: {len(role_type._tableContexts)}")

        # Enable table rendering with proper support
        model_xbrl._hasTableRendering = True
        model_xbrl._first_table_linkrole = first_table_linkrole
        
        # Verify table setup
        log_debug("\nVerifying table setup:")
        log_debug(f"Financial statements found: {len(financial_statements)}")
        log_debug(f"Table facts set up: {len(model_xbrl._tableFacts)}")
        log_debug(f"First table link role: {first_table_linkrole}")
        
        return first_table_linkrole

    except Exception as e:
        log_debug(f"Error in table structure enhancement: {str(e)}", include_trace=True)
        model_xbrl._hasTableRendering = False
        return None

def get_access_statistics(self) -> dict:
    """Get comprehensive access statistics"""
    stats = {
        'total_accesses': sum(len(patterns) for patterns in self.access_patterns.values()),
        'conflicts': len(self.conflicts),
        'concepts': {},
        'access_types': defaultdict(int)
    }
    
    for concept, accesses in self.access_patterns.items():
        stats['concepts'][concept] = {
            'total_accesses': len(accesses),
            'access_types': Counter(a['type'] for a in accesses),
            'has_conflicts': any(c['concept'] == concept for c in self.conflicts)
        }
        
        for access in accesses:
            stats['access_types'][access['type']] += 1
            
    return stats

    # working
    
# Main hub of this plugin 
def filing_loaded(model_xbrl, *args, **kwargs):
    if not model_xbrl:
        return
    # measure total execution time
    start_time = perf_counter()  # Add this at the start        
    
    logger.initialize(model_xbrl)
    
    # Get settings from config file
    # Initialize config handler first
    model_xbrl._config_handler = ConfigHandler(os.path.dirname(__file__))
    print("\nConfig handler initialized")
    logger.initialize(model_xbrl)    
    # print("\n=== File Export Configuration ===")
    # print(f"Filing URI: {model_xbrl.uri}")
    # print(f"Is Online Filing: {model_xbrl.uri.startswith(('http://', 'https://'))}")
    # print(f"Export Mode: {model_xbrl._config_handler.config['excel_export']['local_files' if not model_xbrl.uri.startswith(('http://', 'https://')) else 'online_files']['mode']}")
    # print(f"Export Path: {model_xbrl._config_handler.get_export_path(model_xbrl)}")
    # print(f"Logging Enabled: {model_xbrl._config_handler.is_logging_enabled()}")    
    
    
    try:
        filename = os.path.basename(model_xbrl.uri) if model_xbrl.uri else "unknown_filing"
        log_debug(f"\n=== ARC Plugin: Processing {filename} ===")
        
        # Initialize filing analyzer first
        analyzer = FilingAnalyzer(model_xbrl)
        log_debug(f"\nCreated FilingAnalyzer instance: {id(analyzer)}")
        
        # Initialize fact provider with our analyzer instance
        fact_provider = FactProvider(model_xbrl, existing_analyzer=analyzer)
        model_xbrl._fact_provider = fact_provider
        log_debug(f"\nFilingAnalyzer instance in fact_provider: {id(fact_provider.filing_analyzer)}")

        # Create exporter early since we need it for statement processing
        model_xbrl._exporter = ExcelExporter(model_xbrl)
        
        # Enhance table structure and get first linkrole
        try:
            first_table = enhance_table_structure(model_xbrl)
            model_xbrl._first_table_linkrole = first_table
        except Exception as e:
            log_debug(f"Error enhancing table structure: {str(e)}", include_trace=True)
            model_xbrl._first_table_linkrole = None
            
        # Process statements first to allow proper merging
        log_debug("\nProcessing statements for merging...")
        statements = model_xbrl._exporter.get_financial_statements()
        if statements:
            log_debug(f"Found {len(statements)} statements to merge")
            merged_statements = model_xbrl._exporter.merge_split_statements(statements)
            log_debug(f"Merged into {len(merged_statements)} statements")
            
        # Now initialize contexts with merged statements
        log_debug("\nInitializing contexts in FilingAnalyzer")
        analyzer.initialize_contexts()
        
        # Initialize contexts in financial data manager
        log_debug("\nInitializing contexts in FinancialDataManager")
        fact_provider.financial_data_manager.initialize_contexts()
        
        # Store financial data manager reference
        model_xbrl._financial_data_manager = fact_provider.financial_data_manager
        
        # CRITICAL: Ensure model knows it has table rendering before view refresh
        model_xbrl._hasTableRendering = True
        
        # Force table recognition in Arelle's view system
        if hasattr(model_xbrl, 'viewModelObject'):
            log_debug("\nUpdating model view...")
            if model_xbrl._first_table_linkrole:
                model_xbrl.viewModelObject(model_xbrl._first_table_linkrole)
                
        # Status update after view setup
        if hasattr(model_xbrl.modelManager, 'showStatus'):
            model_xbrl.modelManager.showStatus("Table structure enhanced")
        
        # NEW: Explicit table setup verification before view refresh
        log_debug("\nVerifying table setup before view refresh:")
        for roleURI in model_xbrl._tableFacts.keys():
            roleTypes = model_xbrl.roleTypes.get(roleURI, [])
            # if roleTypes:
            #     roleType = roleTypes[0]
            #     log_debug(f"Role: {roleType.definition}")
            #     log_debug(f"- Table Code: {getattr(roleType, '_tableCode', 'None')}")
            #     log_debug(f"- View Properties: {hasattr(roleType, '_viewProperties')}")
            #     log_debug(f"- Table Facts: {len(model_xbrl._tableFacts[roleURI])}")
        
        # Force view refresh after verification
        cntlr = model_xbrl.modelManager.cntlr
        if hasattr(cntlr, 'reloadViews'):
            log_debug("\nReloading views...")
            cntlr.reloadViews(model_xbrl)

        # Run tests if in testing mode
        if getattr(model_xbrl, 'test_mode', False):
            tester = PluginTester(model_xbrl)
            tester.run_all_tests()

        # Now export to Excel
        log_debug("\nStarting Excel export...")
        model_xbrl._exporter.export_tables_to_excel()

                                
        log_debug("=== Processing Complete ===\n")
        log_debug(f"Total execution time: {perf_counter() - start_time:.2f} seconds")  # Add this at the end
        
        print("\n=== File Export Configuration ===")
        print(f"Filing URI: {model_xbrl.uri}")
        print(f"Is Online Filing: {model_xbrl.uri.startswith(('http://', 'https://'))}")
        print(f"Export Mode: {model_xbrl._config_handler.config['excel_export']['local_files' if not model_xbrl.uri.startswith(('http://', 'https://')) else 'online_files']['mode']}")
        print(f"Export Path: {model_xbrl._config_handler.get_export_path(model_xbrl)}")
        print(f"Logging Enabled: {model_xbrl._config_handler.is_logging_enabled()}")            
        
    except Exception as e:
        log_debug(f"\nError in filing_loaded: {str(e)}")
        log_debug(traceback.format_exc())
        model_xbrl.error(
            "ARC:processingError",
            "Error processing filing: %(error)s",
            error=str(e)
        )      
        
def refresh_views(model_xbrl):
    """Refresh all views to reflect changes"""
    cntlr = model_xbrl.modelManager.cntlr
    
    # Update existing views
    for view in model_xbrl.views:
        if hasattr(view, 'view'):
            view.view(model_xbrl)
            
    # Reload controller views if available
    if hasattr(cntlr, 'reloadViews'):
        cntlr.reloadViews(model_xbrl)

def menu_command(cntlr):
    """Enhanced menu command with configuration dialog"""
    try:
        if not hasattr(cntlr, 'modelManager') or not cntlr.modelManager:
            messagebox.showwarning("Warning", "No filing loaded - some options may be limited")
            
        # Initialize config handler
        config_handler = None
            
        # Try to get existing config handler from loaded model
        if hasattr(cntlr.modelManager, 'modelXbrl') and cntlr.modelManager.modelXbrl:
            config_handler = getattr(cntlr.modelManager.modelXbrl, '_config_handler', None)
        
        # Create new config handler if none exists
        if config_handler is None:
            config_handler = ConfigHandler(os.path.dirname(__file__))
            
        # Show configuration dialog
        dialog = ConfigDialog(cntlr.parent, config_handler)
        cntlr.parent.wait_window(dialog)
        
        # If config was updated and we have a model loaded, reinitialize
        if dialog.result and cntlr.modelManager.modelXbrl:
            cntlr.modelManager.modelXbrl._config_handler = config_handler
            filing_loaded(cntlr.modelManager.modelXbrl)
            
    except Exception as e:
        messagebox.showerror("Error", f"Error in menu command: {str(e)}")
        
        
def menu_extend(cntlr, menu, *args, **kwargs):
    """Extend the Arelle menu with configuration option"""
    menu.add_command(
        label="Configure Automated Ratio Calculator Plugin",
        underline=0,
        command=lambda: menu_command(cntlr)
    )

__pluginInfo__ = {
    'name': 'Automated Ratio Calculator',
    'version': '1.0',
    'description': 'Automated Ratio Calculator and enhanced table detection for German and IFRS filings',
    'license': 'Apache-2',
    'author': 'Jahn Ruehrig',
    'copyright': '(c) Copyright 2025',
    'ModelXbrl.LoadComplete': filing_loaded,
    'CntlrWinMain.Menu.Tools': menu_extend,
    'ViewFiling.View': refresh_views,
}

