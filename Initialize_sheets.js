/*
Transactions: Header Row - 2
Invoice Number,	Invoice Date,	Company ID,	Company Name,	Subscription ID,	Payment Cycle,	Amount,	Recipient,	Item Name,	Merchant GST,	Amount Text,	Description,	Discount Amount,	PDF Link,	Email Sent,	Payment Status

Invoice Number => =CONCAT("EDLR/INV/",Row()-2)
Recipient => =VLOOKUP($C3,'Contact Info'!$A$2:$G$1001,7,False)
Item Name => =VLOOKUP($E3,'Subscription Details'!$A$2:$J$1001,10,False)
Merchant GSt => =VLOOKUP($C3,'Company Details'!$A$2:$G$1001,5,False)
Amount Text => =MONEYTEXT(G3,"INR")
Amount Text Alt => 

=IF(OR(LEN(FLOOR(E3,1))>=13,FLOOR(E3,1)<=0),"Out of range",PROPER(SUBSTITUTE(CONCATENATE(CHOOSE(MID(TEXT(INT(E3),REPT(0,12)),1,1)+1,"","one hundred ","two hundred ","three hundred ","four hundred ","five hundred ","six hundred ","seven hundred ","eight hundred ","nine hundred "),CHOOSE(MID(TEXT(INT(E3),REPT(0,12)),2,1)+1,"",CHOOSE(MID(TEXT(INT(E3),REPT(0,12)),3,1)+1,"ten","eleven","twelve","thirteen","fourteen","fifteen","sixteen","seventeen","eighteen","nineteen"),"twenty","thirty","forty","fifty","sixty","seventy","eighty","ninety"),IF(VALUE(MID(TEXT(INT(E3),REPT(0,12)),2,1))>1,CHOOSE(MID(TEXT(INT(E3),REPT(0,12)),3,1)+1,"","-one","-two","-three","-four","-five","-six","-seven","-eight","-nine"),IF(VALUE(MID(TEXT(INT(E3),REPT(0,12)),2,1))=0,CHOOSE(MID(TEXT(INT(E3),REPT(0,12)),3,1)+1,"","one","two","three","four","five","six","seven","eight","nine"),"")),IF(E3>=10^9," billion ",""),CHOOSE(MID(TEXT(INT(E3),REPT(0,12)),4,1)+1,"","one hundred ","two hundred ","three hundred ","four hundred ","five hundred ","six hundred ","seven hundred ","eight hundred ","nine hundred "),CHOOSE(MID(TEXT(INT(E3),REPT(0,12)),5,1)+1,"",CHOOSE(MID(TEXT(INT(E3),REPT(0,12)),6,1)+1,"ten","eleven","twelve","thirteen","fourteen","fifteen","sixteen","seventeen","eighteen","nineteen"),"twenty","thirty","forty","fifty","sixty","seventy","eighty","ninety"),IF(VALUE(MID(TEXT(INT(E3),REPT(0,12)),5,1))>1,CHOOSE(MID(TEXT(INT(E3),REPT(0,12)),6,1)+1,"","-one","-two","-three","-four","-five","-six","-seven","-eight","-nine"),IF(VALUE(MID(TEXT(INT(E3),REPT(0,12)),5,1))=0,CHOOSE(MID(TEXT(INT(E3),REPT(0,12)),6,1)+1,"","one","two","three","four","five","six","seven","eight","nine"),"")),IF(VALUE(MID(TEXT(INT(E3),REPT(0,12)),4,3))>0," million ",""),CHOOSE(MID(TEXT(INT(E3),REPT(0,12)),7,1)+1,"","one hundred ","two hundred ","three hundred ","four hundred ","five hundred ","six hundred ","seven hundred ","eight hundred ","nine hundred "),CHOOSE(MID(TEXT(INT(E3),REPT(0,12)),8,1)+1,"",CHOOSE(MID(TEXT(INT(E3),REPT(0,12)),9,1)+1,"ten","eleven","twelve","thirteen","fourteen","fifteen","sixteen","seventeen","eighteen","nineteen"),"twenty","thirty","forty","fifty","sixty","seventy","eighty","ninety"),IF(VALUE(MID(TEXT(INT(E3),REPT(0,12)),8,1))>1,CHOOSE(MID(TEXT(INT(E3),REPT(0,12)),9,1)+1,"","-one","-two","-three","-four","-five","-six","-seven","-eight","-nine"),IF(VALUE(MID(TEXT(INT(E3),REPT(0,12)),8,1))=0,CHOOSE(MID(TEXT(INT(E3),REPT(0,12)),9,1)+1,"","one","two","three","four","five","six","seven","eight","nine"),"")),IF(VALUE(MID(TEXT(INT(E3),REPT(0,12)),7,3))," thousand ",""),CHOOSE(MID(TEXT(INT(E3),REPT(0,12)),10,1)+1,"","one hundred ","two hundred ","three hundred ","four hundred ","five hundred ","six hundred ","seven hundred ","eight hundred ","nine hundred "),CHOOSE(MID(TEXT(INT(E3),REPT(0,12)),11,1)+1,"",CHOOSE(MID(TEXT(INT(E3),REPT(0,12)),12,1)+1,"ten","eleven","twelve","thirteen","fourteen","fifteen","sixteen","seventeen","eighteen","nineteen"),"twenty","thirty","forty","fifty","sixty","seventy","eighty","ninety"),IF(VALUE(MID(TEXT(INT(E3),REPT(0,12)),11,1))>1,CHOOSE(MID(TEXT(INT(E3),REPT(0,12)),12,1)+1,"","-one","-two","-three","-four","-five","-six","-seven","-eight","-nine"),IF(VALUE(MID(TEXT(INT(E3),REPT(0,12)),11,1))=0,CHOOSE(MID(TEXT(INT(E3),REPT(0,12)),12,1)+1,"","one","two","three","four","five","six","seven","eight","nine"),""))),"  "," ")&IF(FLOOR(E3,1)>1," Rupees"," Rupee")))

Description => =VLOOKUP($E3,'Subscription Details'!$A$2:$J$1001,4,False)

Company Details: Header Row - 1
Unique ID,	Company Name,	Address,	State,	GST Number,	No. of Outlets,	Status,	Discount Type,	Discount Cycle,	Discount Amount

Unique ID => =CONCAT("EDLR",Row()-1)
State => =LEFT(E2,2)

Contact Info: Header Row - 1
Company ID,	POC,	Company,	Name,	Designation,	Contact Number,	Email

Company => =VLOOKUP($A2,'Company Details'!$A$2:$B$1001,2,False)

Subscription Details: Header Row - 1
Subscription ID,	Name,	Company ID,	Subscription Details,	Start Date,	Upcoming Invoice Date,	Subscription Amount,	Scope,	Payment Cycle,	Item Name

Subscription ID => =CONCAT("EDLR/SUB/",Row()-1)
Name => =VLOOKUP($C2,'Company Details'!$A$2:$B$1001,2,False)
Upcoming Invoice Date => =IF($I2="Monthly",EOMONTH(NOW(),0)+1,IF($I2="Quarterly",DATE(IF(ROUNDUP(MONTH(NOW())/3)>3,YEAR(NOW())+1,YEAR(NOW())),IF(((ROUNDUP(MONTH(NOW())/3)*3)+1)>12,1,((ROUNDUP(MONTH(NOW())/3)*3)+1)),1),IF($I2="Semiannually",DATE(IF(MONTH(NOW())<6,YEAR(NOW()),YEAR(NOW())+1),IF(MONTH(NOW())<6,7,1),1),IF($I2="Annually",DATE(YEAR(NOW())+1,1,1),"Error"))))




Settings

*/
function myFunction() {

}
