# ExceltoQBOConverter
Excel VBA formatter that converts all excel files (in macro enabled workbooks) into QBO files (works for Quickbooks 2019, untested on earlier versions and may potentially work) 

QBO files are used in Quickbooks to import bank data and help reconcile bank transactions 

------------

To use: 
1.) Take the date column in your spreadsheet and use the following formula in conversion: =TEXT(Cell, "YYYYMMDD") and convert all dates into the format listed. 
2.) Generate your FIT ID (which follows the following format: YYYYMMDD##) and generate it for each row in your spreadsheet 
2.) Copy this to the beginning of your QBO file (you can use your IDE or editor to open it up - like Visual Studio Code) and add your information to the fields: 
OFXHEADER:100
DATA:OFXSGML
VERSION:102
SECURITY:NONE
ENCODING:USASCII
CHARSET:1252
COMPRESSION:NONE
OLDFILEUID:NONE
NEWFILEUID:NONE

<OFX>
 <SIGNONMSGSRSV1>
  <SONRS>
   <STATUS>
    <CODE>0
    <SEVERITY>INFO
   </STATUS>
   <DTSERVER>20190216140557
   <LANGUAGE>ENG
   <INTU.BID>2430
  </SONRS>
 </SIGNONMSGSRSV1>
 <BANKMSGSRSV1>
  <STMTTRNRS>
   <TRNUID>0
   <STATUS>
    <CODE>0
    <SEVERITY>INFO
   </STATUS>
   <STMTRS>
    <CURDEF>USD
    <BANKACCTFROM>
     <BANKID>bank #
     <ACCTID>Acct #
     <ACCTTYPE>CHECKING
    </BANKACCTFROM>
    <BANKTRANLIST>
     <DTSTART>Date Start (first transaction) (YYYYMMDD)
     <DTEND>Date End (last transaction) (YYYYMMDD)
       
3.) Copy this to the end of your QBO File: 

</BANKTRANLIST>
    <LEDGERBAL>
     <BALAMT>0
     <DTASOF>Date End (last transaction) (YYYYMMDD)
    </LEDGERBAL>
   </STMTRS>
  </STMTTRNRS>
 </BANKMSGSRSV1>
</OFX>


4.) Use the excel macro to generate the data once it is in the correct format and place the data in between the Header and the Footer 

5.) Also, refer to this documentation in case your fields have any invalid characters for QBO files: https://quickbooks.intuit.com/community/Help-Articles/Message-quot-No-new-transactions-quot-when-importing-web-connect/m-p/216067


