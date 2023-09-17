//Date Logic
var currentDate = new Date();
var year = currentDate.getFullYear();
var month = (currentDate.getMonth() + 1 < 10) ? '0' + (currentDate.getMonth() + 1) : (currentDate.getMonth() + 1);
var day = (currentDate.getDate() < 10) ? '0' + currentDate.getDate() : currentDate.getDate();

var systemDate = `${year}-${month}-${day}`;

global.strGlobalDate = systemDate;

//Default Variables -- No need to change//
global.strDefaultPathIN = "C:/Users/TIRUPRU/OneDrive - C.H. Robinson/TIRUPRU/Application Folder/Public";
global.strDefaultPathINExcelAndXML = "C:/Users/TIRUPRU/OneDrive - C.H. Robinson/TIRUPRU/Application Folder/Templates/ExcelSheetsAndXML";
//global.strDefaultPathXMLOut = "C:/Users/TIRUPRU/OneDrive - C.H. Robinson/TIRUPRU/Prudhvi Workspace/Project XML Folder";
global.strDefaultCarrierPathExcelDSOut = "C:/Users/TIRUPRU/OneDrive - C.H. Robinson/TIRUPRU/Application Folder/Result Excel Data Sheet/Carrier";
global.strDefaultCustPathExcelDSOut = "C:/Users/TIRUPRU/OneDrive - C.H. Robinson/TIRUPRU/Application Folder/Result Excel Data Sheet/Customer";
global.strDefaultPathExcelCarrierLodestarOut = "C:/Users/TIRUPRU/OneDrive - C.H. Robinson/TIRUPRU/Application Folder/Result Excel Lodestars/Carrier";
global.strDefaultPathExcelCustLodestarOut = "C:/Users/TIRUPRU/OneDrive - C.H. Robinson/TIRUPRU/Application Folder/Result Excel Lodestars/Customer";
//global.strDefaultTEDIPathXMLOut = "//gis_si/Nonprod_SI/TRN/AutomatedImporter/In";
global.strDefaultCarrierPathXMLOut = "C:/Users/TIRUPRU/OneDrive - C.H. Robinson/TIRUPRU/Application Folder/Result XML/Carrier";
global.strDefaultCustPathXMLOut = "C:/Users/TIRUPRU/OneDrive - C.H. Robinson/TIRUPRU/Application Folder/Result XML/Customer";


global.strTEDI = "Y";           //Say If Yes - "Y" and If No - "N"
global.strDefault7Letter = "TIRUPRU";
global.strGlobalDevName  = "Prudhvi Tirunagari";
