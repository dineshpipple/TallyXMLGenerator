using System.Xml.Serialization;

// XmlSerializer serializer = new XmlSerializer(typeof(VOUCHER));
// using (StringReader reader = new StringReader(xml))
// {
//    var test = (VOUCHER)serializer.Deserialize(reader);
// }


public class LEDGERENTRIESLIST
{
    public LEDGERENTRIESLIST()
    {
        BASICRATEOFINVOICETAXLIST = new BASICRATEOFINVOICETAXLIST();
    }

    [XmlElement(ElementName = "LEDGERNAME")]
    public string LEDGERNAME { get; set; }

    [XmlElement(ElementName = "ISDEEMEDPOSITIVE")]
    public string ISDEEMEDPOSITIVE { get; set; }

    [XmlElement(ElementName = "AMOUNT")]
    public string AMOUNT { get; set; }

    [XmlElement(ElementName = "BASICRATEOFINVOICETAX.LIST")]
    public BASICRATEOFINVOICETAXLIST BASICRATEOFINVOICETAXLIST { get; set; }

    [XmlElement(ElementName = "VATEXPAMOUNT")]
    public double VATEXPAMOUNT { get; set; }

    [XmlElement(ElementName = "MyProperty")]
    public MyProperty MyProperty { get; set; }

    
}

public class MyProperty
{
    public int Test { get; set; }
}

[XmlRoot(ElementName = "ACCOUNTINGALLOCATIONS.LIST")]
public class ACCOUNTINGALLOCATIONSLIST
{

    public ACCOUNTINGALLOCATIONSLIST()
    {

    }

    [XmlElement(ElementName = "LEDGERNAME")]
    public string LEDGERNAME { get; set; }

    [XmlElement(ElementName = "ISDEEMEDPOSITIVE")]
    public string ISDEEMEDPOSITIVE { get; set; }

    [XmlElement(ElementName = "AMOUNT")]
    public string AMOUNT { get; set; }


}

[XmlRoot(ElementName = "ALLINVENTORYENTRIES.LIST")]
public class ALLINVENTORYENTRIESLIST
{
    public ALLINVENTORYENTRIESLIST()
    {
        ACCOUNTINGALLOCATIONSLIST = new ACCOUNTINGALLOCATIONSLIST();
    }

    [XmlElement(ElementName = "STOCKITEMNAME")]
    public string STOCKITEMNAME { get; set; }

    [XmlElement(ElementName = "ISDEEMEDPOSITIVE")]
    public string ISDEEMEDPOSITIVE { get; set; }

    [XmlElement(ElementName = "RATE")]
    public string RATE { get; set; }

    [XmlElement(ElementName = "AMOUNT")]
    public string AMOUNT { get; set; }

    [XmlElement(ElementName = "ACTUALQTY")]
    public string ACTUALQTY { get; set; }

    [XmlElement(ElementName = "BILLEDQTY")]
    public string BILLEDQTY { get; set; }

    [XmlElement(ElementName = "ACCOUNTINGALLOCATIONS.LIST")]
    public ACCOUNTINGALLOCATIONSLIST ACCOUNTINGALLOCATIONSLIST { get; set; }

    [XmlElement(ElementName = "BATCHALLOCATIONS.LIST")]
    public BATCHALLOCATIONSLIST BATCHALLOCATIONSLIST { get; set; }

}

public class BASICRATEOFINVOICETAXLIST
{


    [XmlElement(ElementName = "BASICRATEOFINVOICETAX")]
    public string BASICRATEOFINVOICETAX { get; set; }

    [XmlAttribute(AttributeName = "Type")]
    public string Type { get; set; }

}


[XmlRoot(ElementName = "VOUCHER")]
public class VOUCHER
{

    public VOUCHER()
    {
        LEDGERENTRIESLIST = new List<LEDGERENTRIESLIST>();
        ALLINVENTORYENTRIESLIST = new ALLINVENTORYENTRIESLIST();
    }

    [XmlElement(ElementName = "DATE")]
    public string DATE { get; set; }

    [XmlElement(ElementName = "PARTYNAME")]
    public string PARTYNAME { get; set; }

    [XmlElement(ElementName = "VOUCHERTYPENAME")]
    public string VOUCHERTYPENAME { get; set; }

    [XmlElement(ElementName = "NARRATION")]
    public string NARRATION { get; set; }

    [XmlElement(ElementName = "VOUCHERNUMBER")]
    public string VOUCHERNUMBER { get; set; }

    [XmlElement(ElementName = "PARTYLEDGERNAME")]
    public string PARTYLEDGERNAME { get; set; }

    [XmlElement(ElementName = "PERSISTEDVIEW")]
    public string PERSISTEDVIEW { get; set; }


    [XmlElement(ElementName = "ISINVOICE")]
    public string ISINVOICE { get; set; }

    [XmlElement(ElementName = "LEDGERENTRIES.LIST")]
    public List<LEDGERENTRIESLIST> LEDGERENTRIESLIST { get; set; }

    [XmlElement(ElementName = "ALLINVENTORYENTRIES.LIST")]
    public ALLINVENTORYENTRIESLIST ALLINVENTORYENTRIESLIST { get; set; }

    [XmlAttribute(AttributeName = "VCHTYPE")]
    public string VCHTYPE { get; set; }

    [XmlAttribute(AttributeName = "ACTION")]
    public string ACTION { get; set; }

    [XmlAttribute(AttributeName = "OBJVIEW")]
    public string OBJVIEW { get; set; }


    
    [XmlElement(ElementName = "GSTREGISTRATIONTYPE")]
    public string GSTREGISTRATIONTYPE { get; set; }

    [XmlElement(ElementName = "STATENAME")]
    public string STATENAME { get; set; }

    [XmlElement(ElementName = "COUNTRYOFRESIDENCE")]
    public string COUNTRYOFRESIDENCE { get; set; }

    [XmlElement(ElementName = "PLACEOFSUPPLY")]
    public string PLACEOFSUPPLY { get; set; }

    [XmlElement(ElementName = "PARTYMAILINGNAME")]
    public string PARTYMAILINGNAME { get; set; }

    [XmlElement(ElementName = "CONSIGNEEMAILINGNAME")]
    public string CONSIGNEEMAILINGNAME { get; set; }

    [XmlElement(ElementName = "CONSIGNEESTATENAME")]
    public string CONSIGNEESTATENAME { get; set; }

    [XmlElement(ElementName = "BASICBASEPARTYNAME")]
    public string BASICBASEPARTYNAME { get; set; }

    [XmlElement(ElementName = "FBTPAYMENTTYPE")]
    public string FBTPAYMENTTYPE { get; set; }


    [XmlElement(ElementName = "BASICBUYERNAME")]
    public string BASICBUYERNAME { get; set; }

    [XmlElement(ElementName = "CONSIGNEECOUNTRYNAME")]
    public string CONSIGNEECOUNTRYNAME { get; set; }

    [XmlElement(ElementName = "VCHENTRYMODE")]
    public string VCHENTRYMODE { get; set; }


}



[XmlRoot(ElementName = "HEADER")]
public class HEADER
{

    [XmlElement(ElementName = "TALLYREQUEST")]
    public string TALLYREQUEST { get; set; }
}

[XmlRoot(ElementName = "STATICVARIABLES")]
public class STATICVARIABLES
{

    [XmlElement(ElementName = "SVCURRENTCOMPANY")]
    public string SVCURRENTCOMPANY { get; set; }
}

[XmlRoot(ElementName = "REQUESTDESC")]
public class REQUESTDESC
{
    public REQUESTDESC()
    {
        STATICVARIABLES = new STATICVARIABLES();
    }

    [XmlElement(ElementName = "REPORTNAME")]
    public string REPORTNAME { get; set; }

    [XmlElement(ElementName = "STATICVARIABLES")]
    public STATICVARIABLES STATICVARIABLES { get; set; }
}



[XmlRoot(ElementName = "TALLYMESSAGE")]
public class TALLYMESSAGE
{
    public TALLYMESSAGE()
    {
        VOUCHER =  new VOUCHER();
        STOCKITEM = new STOCKITEM();
        UNIT = new UNIT();
    }

    [XmlElement(ElementName = "VOUCHER")]
    public VOUCHER? VOUCHER { get; set; }

    [XmlElement(ElementName = "LEDGER")]
    public LEDGER? LEDGER { get; set; }

    [XmlElement(ElementName = "STOCKITEM")]
    public STOCKITEM? STOCKITEM { get; set; }

    [XmlElement(ElementName = "UNIT")]
    public UNIT? UNIT { get; set; }


    [XmlAttribute("xmlnUDF")]
    public string TallyUDF { get; set; }

}

[XmlRoot(ElementName = "REQUESTDATA")]
public class REQUESTDATA
{
    public REQUESTDATA()
    {
        TALLYMESSAGE = new List<TALLYMESSAGE>();
    }

    [XmlElement(ElementName = "TALLYMESSAGE")]
    public List<TALLYMESSAGE> TALLYMESSAGE { get; set; }
}
[XmlRoot(ElementName = "IMPORTDATA")]
public class IMPORTDATA
{

    public IMPORTDATA()
    {
        REQUESTDESC = new REQUESTDESC();
        REQUESTDATA= new REQUESTDATA();
    }

    [XmlElement(ElementName = "REQUESTDESC")]
    public REQUESTDESC REQUESTDESC { get; set; }

    [XmlElement(ElementName = "REQUESTDATA")]
    public REQUESTDATA REQUESTDATA { get; set; }
}

[XmlRoot(ElementName = "BODY")]
public class BODY
{
    public BODY()
    {
        IMPORTDATA = new IMPORTDATA();
    }

    [XmlElement(ElementName = "IMPORTDATA")]
    public IMPORTDATA IMPORTDATA { get; set; }
}

[XmlRoot(ElementName = "ENVELOPE")]
public class ENVELOPE
{
    public ENVELOPE()
    {
        HEADER = new HEADER();
        BODY = new BODY();
    }
    [XmlElement(ElementName = "HEADER")]
    public HEADER HEADER { get; set; }

    [XmlElement(ElementName = "BODY")]
    public BODY BODY { get; set; }
}



//

//[XmlRoot(ElementName = "NAME.LIST")]
//public class NAMELIST
//{

//    [XmlElement(ElementName = "NAME")]
//    public string NAME { get; set; }

//    [XmlAttribute(AttributeName = "TYPE")]
//    public string TYPE { get; set; }
//}

[XmlRoot(ElementName = "LANGUAGENAME.LIST")]
public class LANGUAGENAMELIST
{
    public LANGUAGENAMELIST()
    {
        NAMELIST = new NAMELIST();  
    }

    [XmlElement(ElementName = "NAME.LIST")]
    public NAMELIST NAMELIST { get; set; }

    [XmlElement(ElementName = "LANGUAGEID")]
    public int LANGUAGEID { get; set; }
}

//[XmlRoot(ElementName = "STOCKITEM")]
//public class STOCKITEM
//{
//    public STOCKITEM()
//    {
//        LANGUAGENAMELIST = new LANGUAGENAMELIST();
//    }
//    [XmlElement(ElementName = "LANGUAGENAME.LIST")]
//    public LANGUAGENAMELIST LANGUAGENAMELIST { get; set; }

//    [XmlElement(ElementName = "BASEUNITS")]
//    public string BASEUNITS { get; set; }

//    [XmlAttribute(AttributeName = "NAME")]
//    public string NAME { get; set; }

//    [XmlAttribute(AttributeName = "RESERVEDNAME")]
//    public string RESERVEDNAME { get; set; }

//}


[XmlRoot(ElementName = "RATEDETAILS.LIST")]
public class RATEDETAILSLIST
{

    [XmlElement(ElementName = "GSTRATEDUTYHEAD")]
    public string GSTRATEDUTYHEAD { get; set; }

    [XmlElement(ElementName = "GSTRATEVALUATIONTYPE")]
    public string GSTRATEVALUATIONTYPE { get; set; }

    [XmlElement(ElementName = "GSTRATE")]
    public string GSTRATE { get; set; }
}

[XmlRoot(ElementName = "STATEWISEDETAILS.LIST")]
public class STATEWISEDETAILSLIST
{

    public STATEWISEDETAILSLIST()
    {
        RATEDETAILSLIST = new List<RATEDETAILSLIST>();
    }

    [XmlElement(ElementName = "STATENAME")]
    public string STATENAME { get; set; }

    [XmlElement(ElementName = "RATEDETAILS.LIST")]
    public List<RATEDETAILSLIST> RATEDETAILSLIST { get; set; }

}

[XmlRoot(ElementName = "GSTDETAILS.LIST")]
public class GSTDETAILSLIST
{
    public GSTDETAILSLIST()
    {
        STATEWISEDETAILSLIST = new STATEWISEDETAILSLIST();
    }

    [XmlElement(ElementName = "APPLICABLEFROM")]
    public string APPLICABLEFROM { get; set; }

    [XmlElement(ElementName = "CALCULATIONTYPE")]
    public string CALCULATIONTYPE { get; set; }

    [XmlElement(ElementName = "HSNCODE")]
    public string HSNCODE { get; set; }

    [XmlElement(ElementName = "TAXABILITY")]
    public string TAXABILITY { get; set; }

    [XmlElement(ElementName = "STATEWISEDETAILS.LIST")]
    public STATEWISEDETAILSLIST STATEWISEDETAILSLIST { get; set; }

    [XmlElement(ElementName = "TEMPGSTDETAILSLABRATES.LIST")]
    public string TEMPGSTDETAILSLABRATESLIST { get; set; }
}

[XmlRoot(ElementName = "NAME.LIST")]
public class NAMELIST
{

    [XmlElement(ElementName = "NAME")]
    public string NAME { get; set; }

    [XmlAttribute(AttributeName = "TYPE")]
    public string TYPE { get; set; }

}

//[XmlRoot(ElementName = "LANGUAGENAME.LIST")]
//public class LANGUAGENAMELIST
//{

//    [XmlElement(ElementName = "NAME.LIST")]
//    public NAMELIST NAMELIST { get; set; }

//    [XmlElement(ElementName = "LANGUAGEID")]
//    public int LANGUAGEID { get; set; }
//}

[XmlRoot(ElementName = "STOCKITEM")]
public class STOCKITEM
{
    public STOCKITEM()
    {
        GSTDETAILSLIST = new GSTDETAILSLIST();
    }

    [XmlElement(ElementName = "BASEUNITS")]
    public string BASEUNITS { get; set; }

    [XmlElement(ElementName = "GSTAPPLICABLE")]
    public string GSTAPPLICABLE { get; set; }

    [XmlElement(ElementName = "GSTTYPEOFSUPPLY")]
    public string GSTTYPEOFSUPPLY { get; set; }

    [XmlElement(ElementName = "VATAPPLICABLE")]
    public string VATAPPLICABLE { get; set; }

    [XmlElement(ElementName = "GSTDETAILS.LIST")]
    public GSTDETAILSLIST GSTDETAILSLIST { get; set; }

    [XmlElement(ElementName = "LANGUAGENAME.LIST")]
    public LANGUAGENAMELIST LANGUAGENAMELIST { get; set; }

    [XmlAttribute(AttributeName = "NAME")]
    public string NAME { get; set; }

    [XmlAttribute(AttributeName = "RESERVEDNAME")]
    public string RESERVEDNAME { get; set; }


}




[XmlRoot(ElementName = "UNIT")]
public class UNIT
{

    [XmlElement(ElementName = "NAME")]
    public string NAME { get; set; }

    [XmlElement(ElementName = "GSTREPUOM")]
    public string GSTREPUOM { get; set; }

    [XmlElement(ElementName = "ISSIMPLEUNIT")]
    public string ISSIMPLEUNIT { get; set; }


    [XmlElement(ElementName = "DECIMALPLACES")]
    public string DECIMALPLACES { get; set; }
    

    [XmlAttribute(AttributeName = "NAME")]
    public string AttributeNAME { get; set; }

}

// using System.Xml.Serialization;
// XmlSerializer serializer = new XmlSerializer(typeof(LEDGER));
// using (StringReader reader = new StringReader(xml))
// {
//    var test = (LEDGER)serializer.Deserialize(reader);
// }



[XmlRoot(ElementName = "LEDGER")]
public class LEDGER
{

    [XmlElement(ElementName = "PARENT")]
    public string PARENT { get; set; }

    [XmlElement(ElementName = "MAILINGNAME.LIST")]
    public MAILINGNAMELIST MAILINGNAMELIST { get; set; }

    [XmlElement(ElementName = "GSTAPPLICABLE")]
    public string GSTAPPLICABLE { get; set; }

    [XmlElement(ElementName = "GSTTYPEOFSUPPLY")]
    public string GSTTYPEOFSUPPLY { get; set; }

    [XmlElement(ElementName = "SERVICECATEGORY")]
    public string SERVICECATEGORY { get; set; }

 
    [XmlElement(ElementName = "RATEOFTAXCALCULATION")]
    public string RATEOFTAXCALCULATION { get; set; }


    [XmlElement(ElementName = "VATAPPLICABLE")]
    public string VATAPPLICABLE { get; set; }

    [XmlElement(ElementName = "LANGUAGENAME.LIST")]
    public LANGUAGENAMELIST LANGUAGENAMELIST { get; set; }

    [XmlAttribute(AttributeName = "NAME")]
    public string NAME { get; set; }

    //[XmlAttribute(AttributeName = "RESERVEDNAME")]
    //public object RESERVEDNAME { get; set; }

    //[XmlText]
    //public string Text { get; set; }

    [XmlElement(ElementName = "PRIORSTATENAME")]
    public string PRIORSTATENAME { get; set; }

    [XmlElement(ElementName = "COUNTRYNAME")]
    public string COUNTRYNAME { get; set; }

    [XmlElement(ElementName = "GSTREGISTRATIONTYPE")]
    public string GSTREGISTRATIONTYPE { get; set; }

    [XmlElement(ElementName = "TAXTYPE")]
    public string TAXTYPE { get; set; }

    [XmlElement(ElementName = "GSTDUTYHEAD")]
    public string GSTDUTYHEAD { get; set; }


    [XmlElement(ElementName = "COUNTRYOFRESIDENCE")]
    public string COUNTRYOFRESIDENCE { get; set; }

    [XmlElement(ElementName = "ISBILLWISEON")]
    public string ISBILLWISEON { get; set; }

    [XmlElement(ElementName = "LEDSTATENAME")]
    public string LEDSTATENAME { get; set; }

}




[XmlRoot(ElementName = "MAILINGNAME.LIST")]
public class MAILINGNAMELIST
{

    [XmlElement(ElementName = "MAILINGNAME")]
    public string MAILINGNAME { get; set; }

    [XmlAttribute(AttributeName = "TYPE")]
    public string TYPE { get; set; }

    [XmlText]
    public string Text { get; set; }
}





[XmlRoot(ElementName = "BATCHALLOCATIONS.LIST")]
public class BATCHALLOCATIONSLIST
{

    [XmlElement(ElementName = "GODOWNNAME")]
    public string GODOWNNAME { get; set; }

    [XmlElement(ElementName = "BATCHNAME")]
    public string BATCHNAME { get; set; }

    [XmlElement(ElementName = "AMOUNT")]
    public string AMOUNT { get; set; }

    [XmlElement(ElementName = "ACTUALQTY")]
    public string ACTUALQTY { get; set; }

    [XmlElement(ElementName = "BILLEDQTY")]
    public string BILLEDQTY { get; set; }
}
