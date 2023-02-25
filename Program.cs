using NPOI.HSSF.UserModel;
using NPOI.SS.Formula;
using NPOI.SS.Formula.PTG;
using NPOI.SS.UserModel;
using NPOI.XSSF.UserModel;
using System.Configuration;
using System.Xml;
using System.Xml.Serialization;

var obj = new CreateClass();



public class RequiredName
{
    public string CompanyName { get; set; }
    public string PartyName { get; set; }
    public string VOUCHERTYPENAME { get; set; }
    public string ACCOUNTINGALLOCATIONS { get; set; }
    public string BASICRATEOFINVOICETAX { get; set; }
    public string GST { get; set; }
    public string SGST { get; set; }
    public string CGST { get; set; }
    public string Unit { get; set; }
}



public class CreateClass
{
    public string Name { get; set; }

    public CreateClass()
    {
        Console.WriteLine("Reading values from settings.");
        var setting = ConfigurationManager.AppSettings;
        var required = new RequiredName()
        {
            CompanyName = setting.GetValues("CompanyName").FirstOrDefault().ToString(),

            PartyName = setting.GetValues("PartyName").FirstOrDefault().ToString(),
            VOUCHERTYPENAME = setting.GetValues("VOUCHERTYPENAME").FirstOrDefault().ToString(),
            ACCOUNTINGALLOCATIONS = setting.GetValues("ACCOUNTINGALLOCATIONS").FirstOrDefault().ToString(),
            BASICRATEOFINVOICETAX = setting.GetValues("BASICRATEOFINVOICETAX").FirstOrDefault().ToString(),
            GST = setting.GetValues("GST").FirstOrDefault().ToString(),
            SGST = setting.GetValues("SGST").FirstOrDefault().ToString(),
            CGST = setting.GetValues("CGST").FirstOrDefault().ToString(),
            Unit = setting.GetValues("Unit").FirstOrDefault().ToString(),


        };
        CreateLedger();


        CreateXMLFromExcel("Data.xlsx", requred);
    }

    public void CreateLedger()
    {
        var tallyLeger = new List<TALLYMESSAGE>();
        tallyLeger.Add(new TALLYMESSAGE()
        {
            STOCKITEM = null,
            TallyUDF = null,
            UNIT = null,
            VOUCHER = null,
            LEDGER = IGSTLeger(ConfigurationManager.AppSettings.GetValues("IGSTName").FirstOrDefault().ToString(),
            ConfigurationManager.AppSettings.GetValues("IGSTRate").FirstOrDefault().ToString())
        });


        tallyLeger.Add(new TALLYMESSAGE()
        {
            STOCKITEM = null,
            TallyUDF = null,
            UNIT = null,
            VOUCHER = null,
            LEDGER = SGSTLeger(ConfigurationManager.AppSettings.GetValues("SGSTName").FirstOrDefault().ToString(),
            ConfigurationManager.AppSettings.GetValues("SGSTRate").FirstOrDefault().ToString())
        });

        tallyLeger.Add(new TALLYMESSAGE()
        {
            STOCKITEM = null,
            TallyUDF = null,
            UNIT = null,
            VOUCHER = null,
            LEDGER = CGSTLeger(ConfigurationManager.AppSettings.GetValues("CGSTName").FirstOrDefault().ToString(),
            ConfigurationManager.AppSettings.GetValues("CGSTRate").FirstOrDefault().ToString())
        });

        tallyLeger.Add(new TALLYMESSAGE()
        {
            STOCKITEM = null,
            TallyUDF = null,
            UNIT = null,
            VOUCHER = null,
            LEDGER = SalesLeger(ConfigurationManager.AppSettings.GetValues("SalesLedger").FirstOrDefault().ToString())
        });
        tallyLeger.Add(new TALLYMESSAGE()
        {
            STOCKITEM = null,
            TallyUDF = null,
            UNIT = null,
            VOUCHER = null,
            LEDGER = PartyLeger(ConfigurationManager.AppSettings.GetValues("PartyLedger").FirstOrDefault().ToString(),
            ConfigurationManager.AppSettings.GetValues("PartyLedgerState").FirstOrDefault().ToString())
        });


        tallyLeger.Add(new TALLYMESSAGE()
        {
            STOCKITEM = null,
            VOUCHER = null,
            UNIT = CreateUnit("PCS")
        });


        var mainTag = MainTag();

        mainTag.BODY.IMPORTDATA.REQUESTDATA.TALLYMESSAGE = tallyLeger;

        var ledgerPath = AppDomain.CurrentDomain.BaseDirectory + "Ledger.xml";
        XmlSerializer serializer = new XmlSerializer(typeof(ENVELOPE));
        var writer = new StreamWriter(ledgerPath);
        serializer.Serialize(writer, mainTag);
        writer.Close();

        string text = File.ReadAllText(ledgerPath);
        text = text.Replace("&amp;#4;", "&#4;");
        File.WriteAllText(ledgerPath, text);


    }

    public ENVELOPE MainTag()
    {
        var envelop = new ENVELOPE();
        envelop.HEADER.TALLYREQUEST = "Import Data";
        envelop.BODY.IMPORTDATA.REQUESTDESC.REPORTNAME = "All Masters";
        envelop.BODY.IMPORTDATA.REQUESTDESC.STATICVARIABLES.SVCURRENTCOMPANY = ConfigurationManager.AppSettings.GetValues("CompanyName").FirstOrDefault().ToString();

        return envelop;
    }


    public void CreateXMLFromExcel(string fileName, RequiredName requiredName)
    {
        ISheet sheet;
        var tallyVoucher = new List<TALLYMESSAGE>();
        var tallyStock = new List<TALLYMESSAGE>();
        var tallyUnit = new List<TALLYMESSAGE>();
        var stockDictionary = new Dictionary<string, IRow>();
        XmlSerializer serializer = new XmlSerializer(typeof(ENVELOPE));


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
                    VOUCHER = CreateVouhcer(row, requiredName, evaluator),
                    STOCKITEM = null,
                    UNIT = null,
                });

                if (stockDictionary.ContainsKey(row.GetCell(3).ToString()) == false)
                {
                    stockDictionary.Add(GetCellValue(evaluator.Evaluate(row.GetCell(3))), row);
                }
            }

            foreach (KeyValuePair<string, IRow> ele1 in stockDictionary)
            {
                tallyStock.Add(new TALLYMESSAGE
                {
                    STOCKITEM = CreateStock(ele1.Value, requiredName, evaluator),
                    VOUCHER = null,
                    UNIT = null,
                });
            }

            var envelop = new ENVELOPE();
            envelop.HEADER.TALLYREQUEST = "Import Data";
            envelop.BODY.IMPORTDATA.REQUESTDESC.REPORTNAME = "All Masters";
            envelop.BODY.IMPORTDATA.REQUESTDESC.STATICVARIABLES.SVCURRENTCOMPANY = requiredName.CompanyName;



            //Stock
            envelop.BODY.IMPORTDATA.REQUESTDATA.TALLYMESSAGE = tallyStock;
            var stockPath = AppDomain.CurrentDomain.BaseDirectory + "Stock.xml";
            var writer = new StreamWriter(stockPath);
            serializer.Serialize(writer, envelop);
            writer.Close();


            Console.WriteLine("Stock.xml file successfully created.");

            //Stock
            envelop.BODY.IMPORTDATA.REQUESTDATA.TALLYMESSAGE = tallyVoucher;
            var salesPath = AppDomain.CurrentDomain.BaseDirectory + "Sales.xml";
            writer = new StreamWriter(salesPath);
            serializer.Serialize(writer, envelop);
            writer.Close();

            Console.WriteLine("Sales.xml file successfully created.");

            tallyUnit.Add(new TALLYMESSAGE()
            {
                STOCKITEM = null,
                VOUCHER = null,
                UNIT = CreateUnit(requiredName)
            });

            //Stock
            envelop.BODY.IMPORTDATA.REQUESTDATA.TALLYMESSAGE = tallyUnit;
            var unitPath = AppDomain.CurrentDomain.BaseDirectory + "Unit.xml";
            writer = new StreamWriter(unitPath);
            serializer.Serialize(writer, envelop);
            writer.Close();
            

            Console.WriteLine("Unit.xml file successfully created.");


        }
    }

    private VOUCHER CreateVouhcer(IRow row, RequiredName requiredName, IFormulaEvaluator evaluator)
    {
        List<TALLYMESSAGE> tallyMessage = new List<TALLYMESSAGE>();
        VOUCHER vouhcer = new VOUCHER();
        var voucher = new VOUCHER();

        voucher.VCHTYPE = requiredName.VOUCHERTYPENAME;
        voucher.VOUCHERTYPENAME = requiredName.VOUCHERTYPENAME;
        voucher.ACTION = "Create";
        voucher.OBJVIEW = "Invoice Voucher View";
        voucher.DATE = Convert.ToDateTime(row.GetCell(0).ToString()).ToString("yyyyMMdd");
        voucher.PARTYLEDGERNAME = requiredName.PartyName;

        voucher.NARRATION = row.GetCell(1).StringCellValue;
        voucher.VOUCHERNUMBER = GetCellValue(evaluator.Evaluate(row.GetCell(2)));
        voucher.PARTYLEDGERNAME = requiredName.PartyName;
        voucher.ISINVOICE = "Yes";

        voucher.LEDGERENTRIESLIST.Add(new LEDGERENTRIESLIST()
        {
            LEDGERNAME = requiredName.PartyName,
            ISDEEMEDPOSITIVE = "Yes",
            AMOUNT = "-" + GetCellValue(evaluator.Evaluate(row.GetCell(10))),
        });

        voucher.ALLINVENTORYENTRIESLIST.STOCKITEMNAME = GetCellValue(evaluator.Evaluate(row.GetCell(3)));
        voucher.ALLINVENTORYENTRIESLIST.ISDEEMEDPOSITIVE = "No";
        voucher.ALLINVENTORYENTRIESLIST.RATE = GetCellValue(evaluator.Evaluate(row.GetCell(5))) + "/" + requiredName.Unit;
        voucher.ALLINVENTORYENTRIESLIST.AMOUNT = GetCellValue(evaluator.Evaluate(row.GetCell(6)));
        voucher.ALLINVENTORYENTRIESLIST.ACTUALQTY = GetCellValue(evaluator.Evaluate(row.GetCell(4))) + " " + requiredName.Unit;
        voucher.ALLINVENTORYENTRIESLIST.BILLEDQTY = GetCellValue(evaluator.Evaluate(row.GetCell(4))) + " " + requiredName.Unit;

        voucher.ALLINVENTORYENTRIESLIST.ACCOUNTINGALLOCATIONSLIST.LEDGERNAME = requiredName.ACCOUNTINGALLOCATIONS;
        voucher.ALLINVENTORYENTRIESLIST.ACCOUNTINGALLOCATIONSLIST.ISDEEMEDPOSITIVE = "No";
        voucher.ALLINVENTORYENTRIESLIST.ACCOUNTINGALLOCATIONSLIST.AMOUNT = GetCellValue(evaluator.Evaluate(row.GetCell(6)));


        if (string.IsNullOrWhiteSpace(requiredName.GST) == false)
        {
            voucher.LEDGERENTRIESLIST.Add(new LEDGERENTRIESLIST()
            {
                LEDGERNAME = requiredName.GST,
                ISDEEMEDPOSITIVE = "No",
                AMOUNT = GetCellValue(evaluator.Evaluate(row.GetCell(7))),
                BASICRATEOFINVOICETAXLIST = new BASICRATEOFINVOICETAXLIST() { BASICRATEOFINVOICETAX = requiredName.BASICRATEOFINVOICETAX, Type = "Number" }
            });
        }

        if (string.IsNullOrWhiteSpace(requiredName.SGST) == false && string.IsNullOrWhiteSpace(requiredName.CGST) == false)
        {
            var halftRate = Convert.ToInt32(requiredName.BASICRATEOFINVOICETAX) / 2;

            voucher.LEDGERENTRIESLIST.Add(new LEDGERENTRIESLIST()
            {
                LEDGERNAME = requiredName.SGST,
                ISDEEMEDPOSITIVE = "No",
                AMOUNT = GetCellValue(evaluator.Evaluate(row.GetCell(8))),
                BASICRATEOFINVOICETAXLIST = new BASICRATEOFINVOICETAXLIST() { BASICRATEOFINVOICETAX = halftRate.ToString(), Type = "Number" }
            });


            voucher.LEDGERENTRIESLIST.Add(new LEDGERENTRIESLIST()
            {
                LEDGERNAME = requiredName.CGST,
                ISDEEMEDPOSITIVE = "No",
                AMOUNT = GetCellValue(evaluator.Evaluate(row.GetCell(9))),
                BASICRATEOFINVOICETAXLIST = new BASICRATEOFINVOICETAXLIST() { BASICRATEOFINVOICETAX = halftRate.ToString(), Type = "Number" }
            });
        }
        return voucher;
    }

    private STOCKITEM CreateStock(IRow row, RequiredName requiredName, IFormulaEvaluator evaluator)
    {
        var rateDetailList = new List<RATEDETAILSLIST>();

        if (string.IsNullOrWhiteSpace(requiredName.SGST) == false && string.IsNullOrWhiteSpace(requiredName.CGST) == false)
        {
            rateDetailList.Add(new RATEDETAILSLIST()
            {
                GSTRATEDUTYHEAD = "Central Tax",
                GSTRATEVALUATIONTYPE = "Based on Value",
                GSTRATE = GetCellValue(evaluator.Evaluate(row.GetCell(9))),
            });

            rateDetailList.Add(new RATEDETAILSLIST()
            {
                GSTRATEDUTYHEAD = "State Tax",
                GSTRATEVALUATIONTYPE = "Based on Value",
                GSTRATE = GetCellValue(evaluator.Evaluate(row.GetCell(8))),
            });
        }
        rateDetailList.Add(new RATEDETAILSLIST()
        {
            GSTRATEDUTYHEAD = "Integrated Tax",
            GSTRATEVALUATIONTYPE = "Based on Value",
            GSTRATE = GetCellValue(evaluator.Evaluate(row.GetCell(7))),
        });

        return new STOCKITEM()
        {
            NAME = Name,
            GSTAPPLICABLE = @"&#4; Applicable",
            VATAPPLICABLE = @"&#4; Applicable",
            GSTTYPEOFSUPPLY = "Goods",
            GSTDETAILSLIST = new GSTDETAILSLIST()
            {
                APPLICABLEFROM = row.GetCell(12).ToString(),
                CALCULATIONTYPE = "On Value",
                HSNCODE = GetCellValue(evaluator.Evaluate(row.GetCell(11))),
                TAXABILITY = "Taxable",
                STATEWISEDETAILSLIST = new STATEWISEDETAILSLIST()
                {
                    STATENAME = GetCellValue(evaluator.Evaluate(row.GetCell(13))),
                    RATEDETAILSLIST = rateDetailList,
                }
            },

            LANGUAGENAMELIST = new LANGUAGENAMELIST()
            {
                LANGUAGEID = 1033,
                NAMELIST = new NAMELIST()
                {
                    NAME = GetCellValue(evaluator.Evaluate(row.GetCell(3))),
                    TYPE = "String"
                }
            },
            BASEUNITS = requiredName.Unit,
            RESERVEDNAME = "",

        };
    }


    private UNIT CreateUnit(RequiredName requiredName)
    {
        var GSTREPUOM = "";
        switch (requiredName.Unit.ToUpper())
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
            NAME = requiredName.Unit,
            AttributeNAME = requiredName.Unit,
            GSTREPUOM = GSTREPUOM,
            ISSIMPLEUNIT = "Yes",
            DECIMALPLACES = "2"
        };
    }



    private UNIT CreateUnit(string unit)
    {
        var GSTREPUOM = "";
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
        LEDGER ledger = new LEDGER();
        ledger.NAME = name + rate + " %";

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
        LEDGER ledger = new LEDGER();
        ledger.NAME = name + rate + " %";
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
        LEDGER ledger = new LEDGER();
        ledger.NAME = name + rate + " %";
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

    public static string XmlUnescape(string escaped)
    {
        XmlDocument doc = new XmlDocument();
        XmlNode node = doc.CreateElement("root");
        node.InnerXml = escaped;
        return node.InnerText;
    }
}

