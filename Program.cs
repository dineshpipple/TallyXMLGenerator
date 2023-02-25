using NPOI.HSSF.UserModel;
using NPOI.SS.Formula;
using NPOI.SS.Formula.PTG;
using NPOI.SS.UserModel;
using NPOI.XSSF.UserModel;
using System.Configuration;
using System.Xml;
using System.Xml.Serialization;

var obj = new CreateClass();

public class CreateClass
{
    public string VTypeName { get; set; } = "Sales";
    public string PartyLedgerName { get; set; }
    public string SalesLedgerName { get; set; }
    public string IGSTName { get; set; }
    public string IGSTRate { get; set; }
    public string CGSTName { get; set; }
    public string CGSTRate { get; set; }
    public string SGSTName { get; set; }
    public string SGSTRate { get; set; }
    public string Unit { get; set; }
    public string VoucherName { get; set; } = "Sales";

    public List<TALLYMESSAGE> Master;
    public List<TALLYMESSAGE> Transaction;


    public CreateClass()
    {
        Console.WriteLine("Reading values from settings.");
        var setting = ConfigurationManager.AppSettings;
        Master = new List<TALLYMESSAGE>();
        Transaction = new List<TALLYMESSAGE>();
        CreateMaster();
        CreateSalesVoucher("Data.xlsx");

        //CreateXMLFromExcel("Data.xlsx", requred);
    }

    public void CreateMaster()
    {
        Master.Add(new TALLYMESSAGE()
        {
            STOCKITEM = null,
            TallyUDF = null,
            UNIT = null,
            VOUCHER = null,
            LEDGER = IGSTLeger(ConfigurationManager.AppSettings.GetValues("IGSTName").FirstOrDefault().ToString(),
            ConfigurationManager.AppSettings.GetValues("IGSTRate").FirstOrDefault().ToString())
        });


        Master.Add(new TALLYMESSAGE()
        {
            STOCKITEM = null,
            TallyUDF = null,
            UNIT = null,
            VOUCHER = null,
            LEDGER = SGSTLeger(ConfigurationManager.AppSettings.GetValues("SGSTName").FirstOrDefault().ToString(),
            ConfigurationManager.AppSettings.GetValues("SGSTRate").FirstOrDefault().ToString())
        });

        Master.Add(new TALLYMESSAGE()
        {
            STOCKITEM = null,
            TallyUDF = null,
            UNIT = null,
            VOUCHER = null,
            LEDGER = CGSTLeger(ConfigurationManager.AppSettings.GetValues("CGSTName").FirstOrDefault().ToString(),
            ConfigurationManager.AppSettings.GetValues("CGSTRate").FirstOrDefault().ToString())
        });

        Master.Add(new TALLYMESSAGE()
        {
            STOCKITEM = null,
            TallyUDF = null,
            UNIT = null,
            VOUCHER = null,
            LEDGER = SalesLeger(ConfigurationManager.AppSettings.GetValues("SalesLedger").FirstOrDefault().ToString())
        });

        Master.Add(new TALLYMESSAGE()
        {
            STOCKITEM = null,
            TallyUDF = null,
            UNIT = null,
            VOUCHER = null,
            LEDGER = PartyLeger(ConfigurationManager.AppSettings.GetValues("PartyLedger").FirstOrDefault().ToString(),
            ConfigurationManager.AppSettings.GetValues("PartyLedgerState").FirstOrDefault().ToString())
        });


        Master.Add(new TALLYMESSAGE()
        {
            STOCKITEM = null,
            VOUCHER = null,
            UNIT = CreateUnit("PCS"),
            LEDGER = null
        });



    }

    public ENVELOPE MainTag()
    {
        var envelop = new ENVELOPE();
        envelop.HEADER.TALLYREQUEST = "Import Data";
        envelop.BODY.IMPORTDATA.REQUESTDESC.REPORTNAME = "All Masters";
        envelop.BODY.IMPORTDATA.REQUESTDESC.STATICVARIABLES.SVCURRENTCOMPANY = ConfigurationManager.AppSettings.GetValues("CompanyName").FirstOrDefault().ToString();

        return envelop;
    }

    public void CreateSalesVoucher(string fileName)
    {
        ISheet sheet;
        var tallyVoucher = new List<TALLYMESSAGE>();
        var stockDictionary = new Dictionary<string, IRow>();

        Console.WriteLine("Reading file.");
        using (var stream = new FileStream(fileName, FileMode.Open))
        {
            stream.Position = 0;
            XSSFWorkbook xssWorkbook = new XSSFWorkbook(stream);

            var evaluator = xssWorkbook.GetCreationHelper().CreateFormulaEvaluator();

            sheet = xssWorkbook.GetSheetAt(0);
            IRow headerRow = sheet.GetRow(0);

            Console.WriteLine("Validating columns names.");

            if (!IsAllColumnExists(headerRow))
            {
                throw new Exception("Columns not present");
            }

            for (int i = 1; i <= sheet.LastRowNum; i++)
            {
                IRow row = sheet.GetRow(i);
                if (row == null) continue;
                tallyVoucher.Add(new TALLYMESSAGE
                {
                    VOUCHER = CreateSalesVouhcer(VTypeName,
                    Convert.ToDateTime(row.GetCell(0).ToString()).ToString("yyyyMMdd"),
                    PartyLedgerName, 
                    GetCellValue(evaluator.Evaluate(row.GetCell(2))), 
                    GetCellValue(evaluator.Evaluate(row.GetCell(6))),
                    GetCellValue(evaluator.Evaluate(row.GetCell(10))),
                    GetCellValue(evaluator.Evaluate(row.GetCell(3))), 
                    GetCellValue(evaluator.Evaluate(row.GetCell(5))),
                    Unit, 
                    GetCellValue(evaluator.Evaluate(row.GetCell(4))), 
                    IGSTName, 
                    GetCellValue(evaluator.Evaluate(row.GetCell(7))), 
                    IGSTRate, 
                    CGSTName, 
                    GetCellValue(evaluator.Evaluate(row.GetCell(9))), 
                    CGSTRate, 
                    SGSTName, 
                    GetCellValue(evaluator.Evaluate(row.GetCell(8))), 
                    SGSTRate, 
                    VoucherName, 
                    SalesLedgerName, 
                    GetCellValue(evaluator.Evaluate(row.GetCell(13)))),
                    STOCKITEM = null,
                    UNIT = null,
                });

                if (stockDictionary.ContainsKey(row.GetCell(3).ToString()) == false)
                {
                    stockDictionary.Add(GetCellValue(evaluator.Evaluate(row.GetCell(3))), row);
                }
            }

            foreach (KeyValuePair<string, IRow> ele in stockDictionary)
            {
                Master.Add(new TALLYMESSAGE
                {
                    STOCKITEM = CreateStock(ele.Key.ToString(), 18, Unit, 
                    Convert.ToInt32( evaluator.Evaluate(((IRow)ele.Value).GetCell(12)).NumberValue), 
                    evaluator.Evaluate(((IRow)ele.Value).GetCell(11)).StringValue),
                    VOUCHER = null,
                    UNIT = null
                });
            }


            var mainMasterTag = MainTag();
            mainMasterTag.BODY.IMPORTDATA.REQUESTDATA.TALLYMESSAGE = Master;

            var ledgerPath = AppDomain.CurrentDomain.BaseDirectory + "Master.xml";
            XmlSerializer serializer = new XmlSerializer(typeof(ENVELOPE));
            var writer = new StreamWriter(ledgerPath);
            serializer.Serialize(writer, mainMasterTag);
            writer.Close();

            string text = File.ReadAllText(ledgerPath);
            text = text.Replace("&amp;#4;", "&#4;");
            File.WriteAllText(ledgerPath, text);

            Console.WriteLine("Master.xml file successfully created.");



            var mainTransactionTag = MainTag();
            mainTransactionTag.BODY.IMPORTDATA.REQUESTDATA.TALLYMESSAGE = tallyVoucher;

            var transactionPath = AppDomain.CurrentDomain.BaseDirectory + "Transaction.xml";
            
            XmlSerializer serializer2 = new XmlSerializer(typeof(ENVELOPE));
            var writer2 = new StreamWriter(transactionPath);
            serializer2.Serialize(writer2, mainTransactionTag);
            writer2.Close();


            Console.WriteLine("Transaction.xml file successfully created.");


        }
    }


    private UNIT CreateUnit(string unit)
    {
        var GSTREPUOM = "";
        Unit = unit;
        switch (unit)
        {
            case "KG":
                GSTREPUOM = "KGS-KILOGRAMS";
                break;
            case "NOS":
                GSTREPUOM = "NOS-NUMBERS";
                break;
            case "PCS":
                GSTREPUOM = "PCS-PIECES";
                break;
            default:
                break;
        }
        return new UNIT()
        {
            NAME = unit,
            AttributeNAME = unit,
            GSTREPUOM = GSTREPUOM,
            ISSIMPLEUNIT = "Yes",
            DECIMALPLACES = "2"
        };
    }

    private bool IsAllColumnExists(IRow row)
    {
        if (row.GetCell(0).ToString()?.ToUpper() != "DATE")
        {
            return false;
        }

        if (row.GetCell(1).ToString()?.ToUpper() != "NARRATION")
        {
            return false;
        }

        if (row.GetCell(2).ToString()?.ToUpper() != "VOUCHERNUMBER")
        {
            return false;
        }

        if (row.GetCell(3).ToString()?.ToUpper() != "STOCKITEMNAME")
        {
            return false;
        }

        if (row.GetCell(4).ToString()?.ToUpper() != "QTY")
        {
            return false;
        }

        if (row.GetCell(5).ToString()?.ToUpper() != "RATE")
        {
            return false;
        }

        if (row.GetCell(6).ToString()?.ToUpper() != "AMOUNT")
        {
            return false;
        }

        if (row.GetCell(7).ToString()?.ToUpper() != "GST 18")
        {
            return false;
        }
        if (row.GetCell(8).ToString()?.ToUpper() != "SGST 9")
        {
            return false;
        }
        if (row.GetCell(9).ToString()?.ToUpper() != "CGST 9")
        {
            return false;
        }

        if (row.GetCell(10).ToString()?.ToUpper() != "TOTAL")
        {
            return false;
        }

        if (row.GetCell(11).ToString()?.ToUpper() != "HSN")
        {
            return false;
        }

        if (row.GetCell(12).ToString()?.ToUpper() != "APPLICABLEFROM")
        {
            return false;
        }
        if (row.GetCell(13).ToString()?.ToUpper() != "STATENAME")
        {
            return false;
        }

        return true;
    }

    private string GetCellValue(CellValue cell)
    {
        object cValue = string.Empty;
        switch (cell.CellType)
        {
            case CellType.Numeric:
                cValue = cell.NumberValue;
                break;
            case CellType.String:
                cValue = cell.StringValue;
                break;
            case CellType.Boolean:
                cValue = cell.BooleanValue;
                break;
            case CellType.Error:
                cValue = cell.ErrorValue;
                break;
            default:
                cValue = "";
                break;
        }
        return cValue.ToString();
    }

    private LEDGER SalesLeger(string name)
    {
        SalesLedgerName = name;
        LEDGER ledger = new LEDGER();
        ledger.NAME = name;

        ledger.GSTAPPLICABLE = "&#4; Applicable";
        ledger.VATAPPLICABLE = "&#4; Applicable";

        ledger.PARENT = "Sales Accounts";
        ledger.GSTTYPEOFSUPPLY = "Goods";
        //ledger.SERVICECATEGORY = "Not Applicable";
        ledger.LANGUAGENAMELIST = new LANGUAGENAMELIST();
        ledger.LANGUAGENAMELIST.LANGUAGEID = 1033;
        ledger.LANGUAGENAMELIST.NAMELIST = new NAMELIST();
        ledger.LANGUAGENAMELIST.NAMELIST.NAME = ledger.NAME;
        ledger.LANGUAGENAMELIST.NAMELIST.TYPE = "String";
        return ledger;
    }

    private LEDGER PartyLeger(string name, string state)
    {
        PartyLedgerName = name;
        LEDGER ledger = new LEDGER();
        ledger.NAME = name;

        ledger.MAILINGNAMELIST = new MAILINGNAMELIST();
        ledger.MAILINGNAMELIST.TYPE = "String";
        ledger.MAILINGNAMELIST.MAILINGNAME = ledger.NAME;
        ledger.PRIORSTATENAME = state;
        ledger.COUNTRYNAME = "India";
        ledger.GSTREGISTRATIONTYPE = "Unregistered";
        ledger.PARENT = "Sundry Debtors";
        ledger.TAXTYPE = "Others";
        ledger.COUNTRYOFRESIDENCE = "India";
        ledger.ISBILLWISEON = "No";
        ledger.LEDSTATENAME = state;

        ledger.LANGUAGENAMELIST = new LANGUAGENAMELIST();
        ledger.LANGUAGENAMELIST.LANGUAGEID = 1033;

        ledger.LANGUAGENAMELIST.NAMELIST = new NAMELIST();
        ledger.LANGUAGENAMELIST.NAMELIST.NAME = ledger.NAME;
        ledger.LANGUAGENAMELIST.NAMELIST.TYPE = "String";

        return ledger;
    }

    private LEDGER CGSTLeger(string name, string rate)
    {
        CGSTName = name + rate + " %";
        CGSTRate = rate;
        LEDGER ledger = new LEDGER();
        ledger.NAME = CGSTName;

        ledger.PARENT = "Duties & Taxes";
        ledger.TAXTYPE = "GST";
        ledger.GSTDUTYHEAD = "Central Tax";
        //ledger.SERVICECATEGORY = "SERVICECATEGORY";
        ledger.RATEOFTAXCALCULATION = rate;

        ledger.LANGUAGENAMELIST = new LANGUAGENAMELIST();
        ledger.LANGUAGENAMELIST.LANGUAGEID = 1033;

        ledger.LANGUAGENAMELIST.NAMELIST = new NAMELIST();
        ledger.LANGUAGENAMELIST.NAMELIST.NAME = ledger.NAME;
        ledger.LANGUAGENAMELIST.NAMELIST.TYPE = "String";

        return ledger;
    }

    private LEDGER SGSTLeger(string name, string rate)
    {
        SGSTName = name + rate + " %";
        SGSTRate = rate;
        LEDGER ledger = new LEDGER();
        ledger.NAME = SGSTName;
        ledger.PARENT = "Duties & Taxes";
        ledger.TAXTYPE = "GST";
        ledger.GSTDUTYHEAD = "State Tax";
        //ledger.SERVICECATEGORY = "SERVICECATEGORY";
        ledger.RATEOFTAXCALCULATION = rate;

        ledger.LANGUAGENAMELIST = new LANGUAGENAMELIST();
        ledger.LANGUAGENAMELIST.LANGUAGEID = 1033;

        ledger.LANGUAGENAMELIST.NAMELIST = new NAMELIST();
        ledger.LANGUAGENAMELIST.NAMELIST.NAME = ledger.NAME;
        ledger.LANGUAGENAMELIST.NAMELIST.TYPE = "String";

        return ledger;
    }

    private LEDGER IGSTLeger(string name, string rate)
    {
        IGSTName = name + rate + " %";
        IGSTRate = rate;
        LEDGER ledger = new LEDGER();
        ledger.NAME = IGSTName;
        ledger.PARENT = "Duties & Taxes";
        ledger.TAXTYPE = "GST";
        ledger.GSTDUTYHEAD = "Integrated Tax";
        //ledger.SERVICECATEGORY = "SERVICECATEGORY";
        ledger.RATEOFTAXCALCULATION = rate;

        ledger.LANGUAGENAMELIST = new LANGUAGENAMELIST();
        ledger.LANGUAGENAMELIST.LANGUAGEID = 1033;

        ledger.LANGUAGENAMELIST.NAMELIST = new NAMELIST();
        ledger.LANGUAGENAMELIST.NAMELIST.NAME = ledger.NAME;
        ledger.LANGUAGENAMELIST.NAMELIST.TYPE = "String";

        return ledger;
    }

    private STOCKITEM CreateStock(string name, float rate, string unit, int applicableYear, string hsn)
    {
        var rateDetailList = new List<RATEDETAILSLIST>();

        rateDetailList.Add(new RATEDETAILSLIST()
        {
            GSTRATEDUTYHEAD = "Central Tax",
            GSTRATEVALUATIONTYPE = "Based on Value",
            GSTRATE = Convert.ToString(rate / 2),
        });

        rateDetailList.Add(new RATEDETAILSLIST()
        {
            GSTRATEDUTYHEAD = "State Tax",
            GSTRATEVALUATIONTYPE = "Based on Value",
            GSTRATE = Convert.ToString(rate / 2),
        });
        rateDetailList.Add(new RATEDETAILSLIST()
        {
            GSTRATEDUTYHEAD = "Integrated Tax",
            GSTRATEVALUATIONTYPE = "Based on Value",
            GSTRATE = Convert.ToString(rate),
        });

        return new STOCKITEM()
        {
            NAME = name,
            GSTAPPLICABLE = @"&#4; Applicable",
            VATAPPLICABLE = @"&#4; Applicable",
            GSTTYPEOFSUPPLY = "Goods",
            BASEUNITS = unit,
            RESERVEDNAME = "",

            GSTDETAILSLIST = new GSTDETAILSLIST()
            {
                APPLICABLEFROM = new DateTime(applicableYear, 4, 1).ToString("yyyyMMdd"),
                CALCULATIONTYPE = "On Value",
                HSNCODE = hsn,
                TAXABILITY = "Taxable",
                STATEWISEDETAILSLIST = new STATEWISEDETAILSLIST()
                {
                    STATENAME = "&#4; Any",
                    RATEDETAILSLIST = rateDetailList,
                }
            },

            LANGUAGENAMELIST = new LANGUAGENAMELIST()
            {
                LANGUAGEID = 1033,
                NAMELIST = new NAMELIST()
                {
                    NAME = name,
                    TYPE = "String"
                }
            },

        };
    }


    private VOUCHER CreateSalesVouhcer(string vhType,string vdate, string partyledgerName, string vnumber, string taxableAmount, string withTaxAmount, string stockItemName, string rate, string unit, string qty, string igstLegderName, string igstLegderNameAmt, string igstLegderNameRate, string cgstLegderName, string cgstLegderNameAmt, string cgstLegderNameRate, string sgstLegderName, string sgstLegderNameAmt, string sgstLegderNameRate, string voucherName = "Sale", string salesVoucher = "Amazon Sale", string partyLedgerState = "Delhi")
    {
        VOUCHER vouhcer = new VOUCHER();
        var voucher = new VOUCHER();

        voucher.VCHTYPE = vhType;
        voucher.VOUCHERTYPENAME = voucherName;
        voucher.ACTION = "Create";
        voucher.OBJVIEW = "Invoice Voucher View";
        voucher.DATE = vdate;
        voucher.PARTYNAME = partyledgerName;

        voucher.PARTYLEDGERNAME = partyledgerName;
        voucher.PARTYMAILINGNAME= partyledgerName;
        voucher.CONSIGNEECOUNTRYNAME = partyledgerName;
        voucher.BASICBASEPARTYNAME = partyledgerName;
        voucher.FBTPAYMENTTYPE = "Default";
        voucher.PERSISTEDVIEW = "Invoice Voucher View";
        voucher.BASICBUYERNAME = partyledgerName;
        voucher.CONSIGNEECOUNTRYNAME = "India";
        voucher.VCHENTRYMODE = "Item Invoice";

        voucher.NARRATION = "Narration";
        voucher.VOUCHERNUMBER = vnumber;
        voucher.GSTREGISTRATIONTYPE = "Unregistered";
        voucher.COUNTRYOFRESIDENCE = "India";
        voucher.PLACEOFSUPPLY = partyLedgerState;
        voucher.CONSIGNEESTATENAME = partyLedgerState;
        voucher.STATENAME= partyLedgerState;



        voucher.ALLINVENTORYENTRIESLIST.STOCKITEMNAME = stockItemName;
        voucher.ALLINVENTORYENTRIESLIST.ISDEEMEDPOSITIVE = "No";
        voucher.ALLINVENTORYENTRIESLIST.RATE = rate + "/" + unit;
        voucher.ALLINVENTORYENTRIESLIST.AMOUNT = taxableAmount;
        voucher.ALLINVENTORYENTRIESLIST.ACTUALQTY = qty + " " + unit;
        voucher.ALLINVENTORYENTRIESLIST.BILLEDQTY = qty + " " + unit;

        voucher.ALLINVENTORYENTRIESLIST.BATCHALLOCATIONSLIST =  new BATCHALLOCATIONSLIST();
        voucher.ALLINVENTORYENTRIESLIST.BATCHALLOCATIONSLIST.GODOWNNAME = "Main Location";
        voucher.ALLINVENTORYENTRIESLIST.BATCHALLOCATIONSLIST.BATCHNAME = "Primary Batch";
        voucher.ALLINVENTORYENTRIESLIST.BATCHALLOCATIONSLIST.AMOUNT = taxableAmount;
        voucher.ALLINVENTORYENTRIESLIST.BATCHALLOCATIONSLIST.ACTUALQTY = qty;
        voucher.ALLINVENTORYENTRIESLIST.BATCHALLOCATIONSLIST.BILLEDQTY= qty;

        voucher.ALLINVENTORYENTRIESLIST.ACCOUNTINGALLOCATIONSLIST.LEDGERNAME = salesVoucher;
        //voucher.ALLINVENTORYENTRIESLIST.ACCOUNTINGALLOCATIONSLIST.ISDEEMEDPOSITIVE = "No";
        voucher.ALLINVENTORYENTRIESLIST.ACCOUNTINGALLOCATIONSLIST.AMOUNT = taxableAmount;


        voucher.LEDGERENTRIESLIST.Add(new LEDGERENTRIESLIST()
        {
            LEDGERNAME = partyledgerName,
            ISDEEMEDPOSITIVE = "Yes",
            AMOUNT = "-" + withTaxAmount,
        });

        switch (partyLedgerState)
        {
            case "Delhi":
                voucher.LEDGERENTRIESLIST.Add(new LEDGERENTRIESLIST()
                {
                    LEDGERNAME = cgstLegderName,
                    ISDEEMEDPOSITIVE = "No",
                    AMOUNT = cgstLegderNameAmt,
                    BASICRATEOFINVOICETAXLIST = new BASICRATEOFINVOICETAXLIST()
                    {
                        BASICRATEOFINVOICETAX = cgstLegderNameRate,
                        Type = "Number"
                    }
                });

                voucher.LEDGERENTRIESLIST.Add(new LEDGERENTRIESLIST()
                {
                    LEDGERNAME = sgstLegderName,
                    ISDEEMEDPOSITIVE = "No",
                    AMOUNT = sgstLegderNameAmt,
                    BASICRATEOFINVOICETAXLIST = new BASICRATEOFINVOICETAXLIST()
                    {
                        BASICRATEOFINVOICETAX = sgstLegderNameRate,
                        Type = "Number"
                    }
                });

                break;
            default:
                voucher.LEDGERENTRIESLIST.Add(new LEDGERENTRIESLIST()
                {
                    LEDGERNAME = igstLegderName,
                    ISDEEMEDPOSITIVE = "No",
                    AMOUNT = igstLegderNameAmt,
                    BASICRATEOFINVOICETAXLIST = new BASICRATEOFINVOICETAXLIST()
                    {
                        BASICRATEOFINVOICETAX = igstLegderNameRate,
                        Type = "Number"
                    }
                });
                break;
        }

        return voucher;
    }

}

