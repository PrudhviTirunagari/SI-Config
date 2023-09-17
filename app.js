var express = require("express");
var app     = express();
var path    = require("path");
var fs = require('fs');
require('C:/Users/TIRUPRU/OneDrive - C.H. Robinson/TIRUPRU/Application Folder/Public/Scripts/ScriptDefaultVar.js');

app.use(express.static(__dirname +"/Public"));
//app.use(express.static(__dirname +"/Scripts"));

app.get('/',function(req,res){
  res.sendFile(path.join(__dirname +'/Public/HTML/first_page.html'));
  console.log("Successfully opened Server");
});

app.post("/openCarrierPage", (req, res) => {    
  res.sendFile(path.join(__dirname +'/Public/HTML/index.html'));
})

app.post("/openCustomerPage", (req, res) => {    
  res.sendFile(path.join(__dirname +'/Public/HTML/customer_page.html'));
})

app.listen(8080);


const bodyParser = require('body-parser'); // Import body-parser module

app.use(express.static(__dirname + "/Public"));
app.use(bodyParser.json()); // Use body-parser middleware

app.post('/submitData', (req, res) => {
  const strLocTCode = req.body.strLocTCode;
  const strLocInpath = req.body.strLocInpath;
  const strLocTNewSetup = req.body.strLocTNewSetup;
  const strLocTLodestar = req.body.strLocTLodestar;
  const strLocCNewSetup = req.body.strLocCNewSetup;
  const strLocCLodestar = req.body.strLocCLodestar;
  const strLocCCode = req.body.strLocCCode;
  const strLocDoc520 = req.body.strLocDoc520;
  const strLocDoc514 = req.body.strLocDoc514;
  const strLocDoc204 = req.body.strLocDoc204;
  const strLocType = req.body.strLocType;
  const strLocstr7Letter = req.body.strLocstr7Letter;
  const strCharReplceProcessOnly = req.body.strCharReplceProcessOnly;
  const strUnicodeProcessOnly = req.body.strUnicodeProcessOnly;
  const str520or514NoMailbox = req.body.str520or514NoMailbox;
  const strTMCEMEA = req.body.strTMCEMEA;
  const stCHRGFS = req.body.stCHRGFS;

  console.log('Received data from client:', strLocTCode, strLocInpath, strLocTNewSetup);

  // Specify the directory path where you want to read and write the file
  var strLocDefaultPathINExcelAndXML = global.strDefaultPathINExcelAndXML;
  var strLocDefaultPathCarrierXMLOut = global.strDefaultCarrierPathXMLOut;
  var strLocDefaultCarrierPathExcelOut = global.strDefaultCarrierPathExcelDSOut;
  var strLocDefaultPathExcelCarrierLodestarOut = global.strDefaultPathExcelCarrierLodestarOut;

  var strLocDefaultPathCustXMLOut = global.strDefaultCustPathXMLOut;
  var strLocDefaultCustPathExcelOut = global.strDefaultCustPathExcelDSOut;
  var strLocDefaultPathExcelCustLodestarOut = global.strDefaultPathExcelCustLodestarOut;


  //var strLocstr7Letter = global.strDefault7Letter;
  var strDevName  = global.strGlobalDevName;
  var strLocDate  = global.strGlobalDate;

  var strLocOBTrans520MapName = "Dummy Map Name";
  var iLocNum = 0;
  var strExt = "";

  if (strLocType == "XML") {
    strExt = ".xml";
  }
  else   if (strLocType == "FF") {
    strExt = ".txt";
  }
  
  var XlsxPopulate = require('xlsx-populate');
  var customDirectory = strLocDefaultPathINExcelAndXML;
  console.log(customDirectory);
  console.log(strLocCLodestar,"strLocCLodestar");   
  // Check if the directory exists, if not, create it
  if (!fs.existsSync(customDirectory)) {
      fs.mkdirSync(customDirectory, { recursive: true });
  }

  if (strLocTCode != "") {
    if (strLocTNewSetup === "Yes") {
      // Construct the file paths
      if (strTMCEMEA === "Yes") {
        var inputFilePath = path.join(customDirectory, 'Template XML MBRR.xml');
        var outputFilePath = path.join(strLocDefaultPathCarrierXMLOut, strLocInpath + "_EMEA[" + strLocstr7Letter + "]" + '.xml');
      }
      else if (stCHRGFS === "Yes") {
        var inputFilePath = path.join(customDirectory, 'Template_GFS FreightTracker XML MBRR.xml');
        var outputFilePath = path.join(strLocDefaultPathCarrierXMLOut, strLocInpath + "_GFS[" + strLocstr7Letter + "]" + '.xml');
      }
      // Read the content of the input file
      fs.readFile(inputFilePath, 'utf8', (err, data) => {
        if (err) {
          console.error('Error: Failed to read the file.', err);
        } else {
          // Replace occurrences of 'strTcode' with the value of strLocTCode
          var updatedContent = data.replace(/strTcode/g, strLocTCode).replace(/strDateTime/g, strLocDate).replace(/strDev_Name/g, strDevName)
          .replace(/strCarrier_Name/g, strLocInpath);

          // Write the updated content to the output file
          fs.writeFile(outputFilePath, updatedContent, 'utf8', (err) => {
            if (err) {
              console.error('Error: Failed to write file to disk.', err);
            } else {
              if (strTMCEMEA === "Yes") {
                console.log(strLocInpath + ' EMEA - Mailbox and Routing rules for Carrier is Created and Dropped');
              }
              else if (stCHRGFS === "Yes") {
                console.log(strLocInpath + ' GFS - Mailbox and Routing rules for Carrier is Created and Dropped');
              }
              //console.log('Full output file path:', outputFilePath); // Log the full file path after writing
            }
          });
        }
      });
    
    if (strTMCEMEA === "Yes") {
      outputExcelNameFilePath = path.join(strLocDefaultCarrierPathExcelOut, strLocInpath + '_EMEA.xlsx');
    }
    else if (stCHRGFS === "Yes") {
      outputExcelNameFilePath = path.join(strLocDefaultCarrierPathExcelOut, strLocInpath + '_GFS.xlsx');
    }

    function applyFormattingToExcel(inputPath, outputPath) {
      XlsxPopulate.fromFileAsync(inputPath).then((workbook) => {
        var sheet = workbook.sheet('Sheet1');

        // Define the specific cell addresses to apply the bold style
        var cellAddresses = ['A1', 'B2', 'C3'];
        cellAddresses.forEach((address) => {
          var cell = sheet.cell(address);
          cell.style({ bold: true });
        });

        // Define the specific substrings to be replaced and their corresponding replacements
        var replacements = [
          { find: 'strTcode', replace: strLocTCode }
        ];

        // Iterate through each cell in the sheet
        sheet.usedRange().forEach((cell) => {
          var cellValue = cell.value();

          if (typeof cellValue === 'string') {
            // Check if the cell value contains any of the substrings to be replaced
            replacements.forEach((replacement) => {
              cell.value(cellValue.replace(new RegExp(replacement.find, 'g'), replacement.replace));
            });
          }
        });

        // Define the specific substrings to be replaced and their corresponding replacements
        var replacements = [
          { find: 'strLoc7Letter', replace: strLocstr7Letter }
      ];

      // Iterate through each cell in the sheet
      sheet.usedRange().forEach((cell) => {
      var cellValue = cell.value();

      if (typeof cellValue === 'string') {
          // Check if the cell value contains any of the substrings to be replaced
          replacements.forEach((replacement) => {
          cell.value(cellValue.replace(new RegExp(replacement.find, 'g'), replacement.replace));
          });
      }
      });	

      // Define the specific substrings to be replaced and their corresponding replacements
      var replacements = [
          { find: 'strDateTime', replace: strLocDate}
      ];

      // Iterate through each cell in the sheet
      sheet.usedRange().forEach((cell) => {
      var cellValue = cell.value();

      if (typeof cellValue === 'string') {
          // Check if the cell value contains any of the substrings to be replaced
          replacements.forEach((replacement) => {
          cell.value(cellValue.replace(new RegExp(replacement.find, 'g'), replacement.replace));
          });
      }
      });	
      
      // Define the specific substrings to be replaced and their corresponding replacements
      var replacements = [,
          { find: 'strOutDIPPath', replace: strLocInpath}
      ];

      // Iterate through each cell in the sheet
      sheet.usedRange().forEach((cell) => {
      var cellValue = cell.value();

      if (typeof cellValue === 'string') {
          // Check if the cell value contains any of the substrings to be replaced
          replacements.forEach((replacement) => {
          cell.value(cellValue.replace(new RegExp(replacement.find, 'g'), replacement.replace));
          });
      }
      });

      // Define the specific substrings to be replaced and their corresponding replacements
      var replacements = [
          { find: 'strInFIPPath', replace: strLocInpath}
      ];

      // Iterate through each cell in the sheet
      sheet.usedRange().forEach((cell) => {
      var cellValue = cell.value();

      if (typeof cellValue === 'string') {
          // Check if the cell value contains any of the substrings to be replaced
          replacements.forEach((replacement) => {
          cell.value(cellValue.replace(new RegExp(replacement.find, 'g'), replacement.replace));
          });
      }
      });	

      if (strTMCEMEA === "Yes"){
        console.log(strLocInpath + ' EMEA - Table Entry sheet is created');
      }
      else if (stCHRGFS === "Yes"){
        console.log(strLocInpath + ' GFS - Table Entry sheet is created');
      }
        // Save the modified workbook
        return workbook.toFileAsync(outputPath);
      });
    }

    if (strTMCEMEA === "Yes") {
      var inputExcelFilePath = path.join(customDirectory, 'Template Table Entry.xlsx');
      var outputExcelFilePath = outputExcelNameFilePath;
    }
    else if (stCHRGFS === "Yes") {
      var inputExcelFilePath = path.join(customDirectory, 'Template_GFS FreightTrackerINT.xlsx');
      var outputExcelFilePath = outputExcelNameFilePath;
    }

    applyFormattingToExcel(inputExcelFilePath, outputExcelFilePath);
  }

        ////Lodestar Creation//
  if (strTMCEMEA === "Yes") {       
      if (strLocTLodestar === "Yes") {
        const outputExcel204LodestarFilePath = path.join(strLocDefaultPathExcelCarrierLodestarOut);
        const output204FilePath = path.join(outputExcel204LodestarFilePath, strLocInpath + "_204XML_Lodestar_V1.xlsx");
        const output214FilePath = path.join(outputExcel204LodestarFilePath, strLocInpath + "_214XML_Lodestar_V1.xlsx");
        var output990FilePath = path.join(outputExcel204LodestarFilePath, strLocInpath + "_990XML_Lodestar_V1.xlsx");
    
        applyFormattingToExcel(customDirectory + '/Template_204_Carrier_Lodestar_V01.xlsx', output204FilePath);
        applyFormattingToExcel(customDirectory + '/Template_214_Carrier_Lodestar_V01.xlsx',output214FilePath);
        applyFormattingToExcel(customDirectory + '/Template_990_Carrier_Lodestar_V01.xlsx', output990FilePath);

      
      //var XlsxPopulate = require('xlsx-populate');

      function applyFormattingToExcel(inputLoadstarPath, outputLoadstarPath) {
        XlsxPopulate.fromFileAsync(inputLoadstarPath).then((workbook) => {
            var sheet = workbook.sheet('Sheet1');
        
            // Define the specific cell addresses to apply the bold style
            var cellAddresses = ['A1', 'B2', 'C3'];
            cellAddresses.forEach((address) => {
            var cell = sheet.cell(address);
            cell.style({ bold: true });
            });

            // Define the specific substrings to be replaced and their corresponding replacements
            var replacements = [
                { find: 'strCarrierName', replace: strLocInpath }
            ];
        
            // Iterate through each cell in the sheet
            sheet.usedRange().forEach((cell) => {
                var cellValue = cell.value();
        
                if (typeof cellValue === 'string') {
                // Check if the cell value contains any of the substrings to be replaced
                replacements.forEach((replacement) => {
                    cell.value(cellValue.replace(new RegExp(replacement.find, 'g'), replacement.replace));
                });
                }
            });         
        
            // Define the specific substrings to be replaced and their corresponding replacements
            var replacements = [
            { find: 'strTcode', replace: strLocTCode }
            ];
        
            // Iterate through each cell in the sheet
            sheet.usedRange().forEach((cell) => {
            var cellValue = cell.value();
        
            if (typeof cellValue === 'string') {
                // Check if the cell value contains any of the substrings to be replaced
                replacements.forEach((replacement) => {
                cell.value(cellValue.replace(new RegExp(replacement.find, 'g'), replacement.replace));
                });
            }
            }); 
        
            // Define the specific substrings to be replaced and their corresponding replacements
            var replacements = [
            { find: 'strLocstr7Letter', replace: strLocstr7Letter }
            ];
        
            // Iterate through each cell in the sheet
            sheet.usedRange().forEach((cell) => {
            var cellValue = cell.value();
        
            if (typeof cellValue === 'string') {
                // Check if the cell value contains any of the substrings to be replaced
                replacements.forEach((replacement) => {
                cell.value(cellValue.replace(new RegExp(replacement.find, 'g'), replacement.replace));
                });
            }
            });	
        
            // Define the specific substrings to be replaced and their corresponding replacements
            var replacements = [
            { find: 'strDateTime', replace: strLocDate}
            ];
        
            // Iterate through each cell in the sheet
            sheet.usedRange().forEach((cell) => {
            var cellValue = cell.value();
        
            if (typeof cellValue === 'string') {
                // Check if the cell value contains any of the substrings to be replaced
                replacements.forEach((replacement) => {
                cell.value(cellValue.replace(new RegExp(replacement.find, 'g'), replacement.replace));
                });
            }
            });	
            
            // Define the specific substrings to be replaced and their corresponding replacements
            var replacements = [
            { find: 'strVLTPath', replace: strLocInpath}
            ];
        
            // Iterate through each cell in the sheet
            sheet.usedRange().forEach((cell) => {
            var cellValue = cell.value();
        
            if (typeof cellValue === 'string') {
                // Check if the cell value contains any of the substrings to be replaced
                replacements.forEach((replacement) => {
                cell.value(cellValue.replace(new RegExp(replacement.find, 'g'), replacement.replace));
                });
            }
            });
        
        
            workbook.toFileAsync(outputLoadstarPath).then(() => {
              console.log(`${outputLoadstarPath} - Lodestar is created`);
            }).catch((err) => {
              console.error('Error: Failed to write file to disk.', err);
            });
          }).catch((err) => {
            console.error('Error: Failed to read the file.', err);
          });
        }
      }
    } 
  }

      //**Customer Code**//
    if (strLocCCode !== "") {
      if (strLocCNewSetup === "Yes") {
        // Construct the file paths

        if (strLocDoc204 === "Yes" && strUnicodeProcessOnly === "Yes") {
            var inputFilePath204 = path.join(customDirectory, '/Template_Customer_IB_UnicodeProcessOnly_XML or FF MBRR.xml');
            var outputFilePath204 = path.join(strLocDefaultPathCustXMLOut, strLocInpath + "_204UniCodeProcessOnlyIB_[" + strLocstr7Letter + "]" + '.xml');
        
            // Read the content of the input file
            fs.readFile(inputFilePath204, 'utf8', (err, data) => {
            if (err) {
                console.error('Error: Failed to read the file.', err);
            } else {
                // Replace occurrences of 'strTcode' with the value of strLocTCode
                var updatedContent = data.replace(/strCcode/g, strLocCCode).replace(/strDateTime/g, strLocDate).replace(/strDev_Name/g, strDevName)
                .replace(/strCustomer_Name/g, strLocInpath);
        
                // Write the updated content to the output file
                fs.writeFile(outputFilePath204, updatedContent, 'utf8', (err) => {
                if (err) {
                    console.error('Error: Failed to write file to disk.', err);
                } else {
                    console.log('File Written Successfully');
                    console.log('Full output file path:', outputFilePath204); // Log the full file path after writing
                }
                });
            }
            });
        }
        else if (strLocDoc204 === "Yes" && strCharReplceProcessOnly === "Yes") {
            var inputFilePath204 = path.join(customDirectory, '/Template_Customer_IB_CharRepalceOnly_XML or FF MBRR.xml');
            var outputFilePath204 = path.join(strLocDefaultPathCustXMLOut, strLocInpath + "_204CharReplaceOnlyIB_[" + strLocstr7Letter + "]" + '.xml');
        
            // Read the content of the input file
            fs.readFile(inputFilePath204, 'utf8', (err, data) => {
            if (err) {
                console.error('Error: Failed to read the file.', err);
            } else {
                // Replace occurrences of 'strTcode' with the value of strLocTCode
                var updatedContent = data.replace(/strCcode/g, strLocCCode).replace(/strDateTime/g, strLocDate).replace(/strDev_Name/g, strDevName)
                .replace(/strCustomer_Name/g, strLocInpath);
        
                // Write the updated content to the output file
                fs.writeFile(outputFilePath204, updatedContent, 'utf8', (err) => {
                if (err) {
                    console.error('Error: Failed to write file to disk.', err);
                } else {
                    console.log('File Written Successfully');
                    console.log('Full output file path:', outputFilePath204); // Log the full file path after writing
                }
                });
            }
            });
        }
        else if (strLocDoc204 === "Yes") {
            var inputFilePath204 = path.join(customDirectory, '/Template_Customer_IB_Normal_XML or FF MBRR.xml');
            var outputFilePath204 = path.join(strLocDefaultPathCustXMLOut, strLocInpath + "_204IB_[" + strLocstr7Letter + "]" + '.xml');
        
            // Read the content of the input file
            fs.readFile(inputFilePath204, 'utf8', (err, data) => {
            if (err) {
                console.error('Error: Failed to read the file.', err);
            } else {
                // Replace occurrences of 'strTcode' with the value of strLocTCode
                var updatedContent = data.replace(/strCcode/g, strLocCCode).replace(/strDateTime/g, strLocDate).replace(/strDev_Name/g, strDevName)
                .replace(/strCustomer_Name/g, strLocInpath);
        
                // Write the updated content to the output file
                fs.writeFile(outputFilePath204, updatedContent, 'utf8', (err) => {
                if (err) {
                    console.error('Error: Failed to write file to disk.', err);
                } else {
                    console.log('File Written Successfully');
                    console.log('Full output file path:', outputFilePath204); // Log the full file path after writing
                }
                });
            }
            });
        }

        if ((strLocDoc520 === "Yes" || strLocDoc514 === "Yes") && str520or514NoMailbox === "") {
            var inputFilePath = path.join(customDirectory, '/Template_Customer_OB_XML or FF MBRR.xml');
            var outputFilePath = path.join(strLocDefaultPathCustXMLOut, strLocInpath + "_520or514OB_[" + strLocstr7Letter + "]" + '.xml');
            
        
            // Read the content of the input file
            fs.readFile(inputFilePath, 'utf8', (err, data) => {
            if (err) {
                console.error('Error: Failed to read the file.', err);
            } else {
                // Replace occurrences of 'strTcode' with the value of strLocTCode
                var updatedContent = data.replace(/strCcode/g, strLocCCode).replace(/strDateTime/g, strLocDate).replace(/strDev_Name/g, strDevName)
                .replace(/strCustomer_Name/g, strLocInpath);
        
                // Write the updated content to the output file
                fs.writeFile(outputFilePath, updatedContent, 'utf8', (err) => {
                if (err) {
                    console.error('Error: Failed to write file to disk.', err);
                } else {
                    console.log('File Written Successfully');
                    console.log('Full output file path:', outputFilePath); // Log the full file path after writing
                }
                });
            }
            });
        }
                     
        const outputExcel204CustTableEntryFilePath = path.join(strLocDefaultCustPathExcelOut);
        let strDocTD204, output204CustFilePath;

        if (strLocDoc204 === "Yes"  && strUnicodeProcessOnly === "Yes") {
            strDocTD204 = "204";
            output204CustFilePath = path.join(outputExcel204CustTableEntryFilePath, strLocInpath + "_" + strDocTD204 + strLocType + "UnicodeProcess_TableEntrySheet.xlsx");
            applyFormattingToExcel(customDirectory + '/Template_Customer Table Entry_IB NAVO_UnicodeProcessOnly_XML or FF.xlsx', output204CustFilePath);
        }
        else if (strLocDoc204 === "Yes"  && strCharReplceProcessOnly === "Yes") {
            strDocTD204 = "204";
            output204CustFilePath = path.join(outputExcel204CustTableEntryFilePath, strLocInpath + "_" + strDocTD204 + strLocType + "CharReplace_TableEntrySheet.xlsx");
            applyFormattingToExcel(customDirectory + '/Template_Customer Table Entry_IB NAVO_CharReplaceOnly_XML or FF.xlsx', output204CustFilePath);
        }
        else if (strLocDoc204 === "Yes") {
          strDocTD204 = "204";
          output204CustFilePath = path.join(outputExcel204CustTableEntryFilePath, strLocInpath + "_" + strDocTD204 + strLocType + "_TableEntrySheet.xlsx");
          applyFormattingToExcel(customDirectory + '/Template_Customer Table Entry_IB NAVO_Normal_XML or FF.xlsx', output204CustFilePath);
        }

        let strDocTD520, output520TDFilePath;
        let strDocTD514, output514TDFilePath;

        if (strLocDoc520 === "Yes" && strLocDoc514 === "Yes") {
            strDocTD520 = "520"; strDocTD514 = "514";
            output520TDFilePath = path.join(outputExcel204CustTableEntryFilePath, strLocInpath + "_" + strDocTD520 + strDocTD514 + "_" + strLocType + "_TableEntrySheet.xlsx");
        } 
        else if (strLocDoc520 === "Yes") {
          strDocTD520 = "520";
          output520TDFilePath = path.join(outputExcel204CustTableEntryFilePath, strLocInpath + "_" + strDocTD520 + strLocType + "_TableEntrySheet.xlsx");
        }          
        else if (strLocDoc514 === "Yes") {
          strDocTD514 = "514";
          output514TDFilePath = path.join(outputExcel204CustTableEntryFilePath, strLocInpath + "_" + strDocTD514 + strLocType + "_TableEntrySheet.xlsx");
        }
    
        
        
        if (strLocDoc514 === "Yes" && strLocDoc520 === "Yes"){
            applyFormattingToExcel(customDirectory + '/Template_Customer Table Entry_ 520 514OB_XML or FF.xlsx', output520TDFilePath);
        }
        else if ( strLocDoc520 === "Yes"){
            applyFormattingToExcel(customDirectory + '/Template_Customer Table Entry_520OB_XML or FF.xlsx',output520TDFilePath);
        } 
        else if ( strLocDoc514 === "Yes"){
            applyFormattingToExcel(customDirectory + '/Template_Customer Table Entry_514OB_XML or FF.xlsx',output514TDFilePath);
        } 


          function applyFormattingToExcel(inputCustExcelTDPath, outputCustExcelTDPath) {
          XlsxPopulate.fromFileAsync(inputCustExcelTDPath).then((workbook) => {
          const sheet = workbook.sheet('Sheet1');

          // Define the specific cell addresses to apply the bold style
          const cellAddresses = ['A1', 'B2', 'C3'];
          cellAddresses.forEach((address) => {
              const cell = sheet.cell(address);
              cell.style({ bold: true });
          });

          // Define the specific substrings to be replaced and their corresponding replacements
          var replacements = [
              { find: 'strCustCCode', replace: strLocCCode }
          ];

          // Iterate through each cell in the sheet
          sheet.usedRange().forEach((cell) => {
              var cellValue = cell.value();

              if (typeof cellValue === 'string') {
              // Check if the cell value contains any of the substrings to be replaced
              replacements.forEach((replacement) => {
                  cell.value(cellValue.replace(new RegExp(replacement.find, 'g'), replacement.replace));
              });
              }
          }); 

          // Define the specific substrings to be replaced and their corresponding replacements
          if (strLocDoc204 === "Yes") {
            var replacements = [
                { find: 'strTransType0', replace: strDocTD204 }
            ];

            // Iterate through each cell in the sheet
            sheet.usedRange().forEach((cell) => {
                var cellValue = cell.value();

                if (typeof cellValue === 'string') {
                // Check if the cell value contains any of the substrings to be replaced
                replacements.forEach((replacement) => {
                    cell.value(cellValue.replace(new RegExp(replacement.find, 'g'), replacement.replace));
                });
                }
            });
          }

          // Define the specific substrings to be replaced and their corresponding replacements
          if (strLocDoc520 === "Yes") {
            var replacements = [
                { find: 'strTransType1', replace: strDocTD520 }
            ];

            // Iterate through each cell in the sheet
            sheet.usedRange().forEach((cell) => {
                var cellValue = cell.value();

                if (typeof cellValue === 'string') {
                // Check if the cell value contains any of the substrings to be replaced
                replacements.forEach((replacement) => {
                    cell.value(cellValue.replace(new RegExp(replacement.find, 'g'), replacement.replace));
                });
                }
            });
          }

          if (strLocDoc514 === "Yes") {
            var replacements = [
                { find: 'strTransType2', replace: strDocTD514 }
            ];

            // Iterate through each cell in the sheet
            sheet.usedRange().forEach((cell) => {
                var cellValue = cell.value();

                if (typeof cellValue === 'string') {
                // Check if the cell value contains any of the substrings to be replaced
                replacements.forEach((replacement) => {
                    cell.value(cellValue.replace(new RegExp(replacement.find, 'g'), replacement.replace));
                });
                }
            });
          }

          // Define the specific substrings to be replaced and their corresponding replacements
          var replacements = [
              { find: 'strTransMapName', replace: strLocOBTrans520MapName }
          ];

          // Iterate through each cell in the sheet
          sheet.usedRange().forEach((cell) => {
              var cellValue = cell.value();

              if (typeof cellValue === 'string') {
              // Check if the cell value contains any of the substrings to be replaced
              replacements.forEach((replacement) => {
                  cell.value(cellValue.replace(new RegExp(replacement.find, 'g'), replacement.replace));
              });
              }
          }); 

          // Define the specific substrings to be replaced and their corresponding replacements
          iLocNum = iLocNum + 1;
          var replacements = [
              { find: 'iNum', replace: iLocNum }
          ];

          // Iterate through each cell in the sheet
          sheet.usedRange().forEach((cell) => {
              var cellValue = cell.value();

              if (typeof cellValue === 'string') {
              // Check if the cell value contains any of the substrings to be replaced
              replacements.forEach((replacement) => {
                  cell.value(cellValue.replace(new RegExp(replacement.find, 'g'), replacement.replace));
              });
              }
          });   
          
          // Define the specific substrings to be replaced and their corresponding replacements
          var replacements = [
              { find: 'strExt', replace: strExt }
          ];

          // Iterate through each cell in the sheet
          sheet.usedRange().forEach((cell) => {
              var cellValue = cell.value();

              if (typeof cellValue === 'string') {
              // Check if the cell value contains any of the substrings to be replaced
              replacements.forEach((replacement) => {
                  cell.value(cellValue.replace(new RegExp(replacement.find, 'g'), replacement.replace));
              });
              }
          });     

          // Define the specific substrings to be replaced and their corresponding replacements
          var replacements = [
              { find: 'str7Letter', replace: strLocstr7Letter }
          ];

          // Iterate through each cell in the sheet
          sheet.usedRange().forEach((cell) => {
              var cellValue = cell.value();

              if (typeof cellValue === 'string') {
              // Check if the cell value contains any of the substrings to be replaced
              replacements.forEach((replacement) => {
                  cell.value(cellValue.replace(new RegExp(replacement.find, 'g'), replacement.replace));
              });
              }
          });	

          // Define the specific substrings to be replaced and their corresponding replacements
          var replacements = [
              { find: 'strDateTime', replace: strLocDate}
          ];

          // Iterate through each cell in the sheet
          sheet.usedRange().forEach((cell) => {
              var cellValue = cell.value();

              if (typeof cellValue === 'string') {
              // Check if the cell value contains any of the substrings to be replaced
              replacements.forEach((replacement) => {
                  cell.value(cellValue.replace(new RegExp(replacement.find, 'g'), replacement.replace));
              });
              }
          });	
          
          // Define the specific substrings to be replaced and their corresponding replacements
          var replacements = [
              { find: 'strCustomerName', replace: strLocInpath}
          ];

          // Iterate through each cell in the sheet
          sheet.usedRange().forEach((cell) => {
              var cellValue = cell.value();

              if (typeof cellValue === 'string') {
              // Check if the cell value contains any of the substrings to be replaced
              replacements.forEach((replacement) => {
                  cell.value(cellValue.replace(new RegExp(replacement.find, 'g'), replacement.replace));
              });
              }
          });

          workbook.toFileAsync(outputCustExcelTDPath).then(() => {
            console.log(`${outputCustExcelTDPath} - Customer Excel Data Sheet is created`);
          }).catch((err) => {
            console.error('Error: Failed to write file to disk.', err);
          });
        }).catch((err) => {
          console.error('Error: Failed to read the file.', err);
        });
      }
    }
           
      //if (strLocOBTransType520 === "520") {
        if (strLocCLodestar === "Yes") {
          const outputExcel204CustLodestarFilePath = path.join(strLocDefaultPathExcelCustLodestarOut);

          let strDoc204, output204CustLSFilePath;
          if (strLocDoc204 === "Yes" && strUnicodeProcessOnly === "Yes") {
            strDoc204 = "204";
            output204CustLSFilePath = path.join(outputExcel204CustLodestarFilePath, strLocInpath + "_" + strDoc204 + strLocType + "UnicodeProcess_Lodestar_V1.xlsx");
            applyFormattingToExcel(customDirectory + '/Template_Customer_IB_XML or FF 204UnicodeProcess only_Lodestar_V01.xlsx', output204CustLSFilePath);
          }
          else if (strLocDoc204 === "Yes" && strCharReplceProcessOnly === "Yes") {
            strDoc204 = "204";
            output204CustLSFilePath = path.join(outputExcel204CustLodestarFilePath, strLocInpath + "_" + strDoc204 + strLocType + "CharReplace_Lodestar_V1.xlsx");
            applyFormattingToExcel(customDirectory + '/Template_Customer_IB_XML or FF 204CharReplace only_Lodestar_V01.xlsx', output204CustLSFilePath);
          }
          else if (strLocDoc204 === "Yes") {
            strDoc204 = "204";
            output204CustLSFilePath = path.join(outputExcel204CustLodestarFilePath, strLocInpath + "_" + strDoc204 + strLocType + "_Lodestar_V1.xlsx");
            applyFormattingToExcel(customDirectory + '/Template_Customer_IB_XML or FF 204_Lodestar_V01.xlsx', output204CustLSFilePath);
          }

          let strDoc520, output520FilePath;
          if (strLocDoc520 === "Yes") {
            strDoc520 = "520";
            output520FilePath = path.join(outputExcel204CustLodestarFilePath, strLocInpath + "_" + strDoc520 + strLocType + "_Lodestar_V1.xlsx");
            applyFormattingToExcel(customDirectory + '/Template_Customer_OB_XML or FF 520_Lodestar_V01.xlsx',output520FilePath);
          }

          let strDoc514, output514FilePath;          
          if (strLocDoc514 === "Yes") {
            strDoc514 = "514";
            output514FilePath = path.join(outputExcel204CustLodestarFilePath, strLocInpath + "_" + strDoc514 + strLocType + "_Lodestar_V1.xlsx");
            applyFormattingToExcel(customDirectory + '/Template_Customer_OB_XML or FF 514_Lodestar_V01.xlsx', output514FilePath);
          }
      
          
        function applyFormattingToExcel(inputCustLodestarPath, outputCustLodestarPath) {
        XlsxPopulate.fromFileAsync(inputCustLodestarPath).then((workbook) => {
        const sheet = workbook.sheet('Sheet1');

        // Define the specific cell addresses to apply the bold style
        const cellAddresses = ['A1', 'B2', 'C3'];
        cellAddresses.forEach((address) => {
            const cell = sheet.cell(address);
            cell.style({ bold: true });
        });

        // Define the specific substrings to be replaced and their corresponding replacements
        var replacements = [
            { find: 'strCustCCode', replace: strLocCCode }
        ];

        // Iterate through each cell in the sheet
        sheet.usedRange().forEach((cell) => {
            var cellValue = cell.value();

            if (typeof cellValue === 'string') {
            // Check if the cell value contains any of the substrings to be replaced
            replacements.forEach((replacement) => {
                cell.value(cellValue.replace(new RegExp(replacement.find, 'g'), replacement.replace));
            });
            }
        }); 

        // Define the specific substrings to be replaced and their corresponding replacements
        if (strLocDoc204 === "Yes") {
            var replacements = [
                { find: 'strTransType0', replace: strDoc204 }
            ];
          }
        if (strLocDoc520 === "Yes") {
          var replacements = [
              { find: 'strTransType1', replace: strDoc520 }
          ];
        }
        if (strLocDoc514 === "Yes") {
          var replacements = [
              { find: 'strTransType2', replace: strDoc514 }
          ];
        }

        // Iterate through each cell in the sheet
        sheet.usedRange().forEach((cell) => {
            var cellValue = cell.value();

            if (typeof cellValue === 'string') {
            // Check if the cell value contains any of the substrings to be replaced
            replacements.forEach((replacement) => {
                cell.value(cellValue.replace(new RegExp(replacement.find, 'g'), replacement.replace));
            });
            }
        });    

        // Define the specific substrings to be replaced and their corresponding replacements
        var replacements = [
            { find: 'strTransMapName', replace: strLocOBTrans520MapName }
        ];

        // Iterate through each cell in the sheet
        sheet.usedRange().forEach((cell) => {
            var cellValue = cell.value();

            if (typeof cellValue === 'string') {
            // Check if the cell value contains any of the substrings to be replaced
            replacements.forEach((replacement) => {
                cell.value(cellValue.replace(new RegExp(replacement.find, 'g'), replacement.replace));
            });
            }
        }); 

        // Define the specific substrings to be replaced and their corresponding replacements
        //iLocNum = iLocNum + 1;
        //var replacements = [
            //{ find: 'iNum', replace: iLocNum }
        //];

        // Iterate through each cell in the sheet
        //sheet.usedRange().forEach((cell) => {
            //var cellValue = cell.value();

            //if (typeof cellValue === 'string') {
            // Check if the cell value contains any of the substrings to be replaced
            //replacements.forEach((replacement) => {
                //cell.value(cellValue.replace(new RegExp(replacement.find, 'g'), replacement.replace));
            //});
            //}
        //});   
        
        // Define the specific substrings to be replaced and their corresponding replacements
        var replacements = [
            { find: 'strExt', replace: strExt }
        ];

        // Iterate through each cell in the sheet
        sheet.usedRange().forEach((cell) => {
            var cellValue = cell.value();

            if (typeof cellValue === 'string') {
            // Check if the cell value contains any of the substrings to be replaced
            replacements.forEach((replacement) => {
                cell.value(cellValue.replace(new RegExp(replacement.find, 'g'), replacement.replace));
            });
            }
        });     

        // Define the specific substrings to be replaced and their corresponding replacements
        var replacements = [
            { find: 'str7Letter', replace: strLocstr7Letter }
        ];

        // Iterate through each cell in the sheet
        sheet.usedRange().forEach((cell) => {
            var cellValue = cell.value();

            if (typeof cellValue === 'string') {
            // Check if the cell value contains any of the substrings to be replaced
            replacements.forEach((replacement) => {
                cell.value(cellValue.replace(new RegExp(replacement.find, 'g'), replacement.replace));
            });
            }
        });	

        // Define the specific substrings to be replaced and their corresponding replacements
        var replacements = [
            { find: 'strDateTime', replace: strLocDate}
        ];

        // Iterate through each cell in the sheet
        sheet.usedRange().forEach((cell) => {
            var cellValue = cell.value();

            if (typeof cellValue === 'string') {
            // Check if the cell value contains any of the substrings to be replaced
            replacements.forEach((replacement) => {
                cell.value(cellValue.replace(new RegExp(replacement.find, 'g'), replacement.replace));
            });
            }
        });	
        
        // Define the specific substrings to be replaced and their corresponding replacements
        var replacements = [
            { find: 'strCustomerName', replace: strLocInpath}
        ];

        // Iterate through each cell in the sheet
        sheet.usedRange().forEach((cell) => {
            var cellValue = cell.value();

            if (typeof cellValue === 'string') {
            // Check if the cell value contains any of the substrings to be replaced
            replacements.forEach((replacement) => {
                cell.value(cellValue.replace(new RegExp(replacement.find, 'g'), replacement.replace));
            });
            }
        });

        workbook.toFileAsync(outputCustLodestarPath).then(() => {
          console.log(`${outputCustLodestarPath} - Customer Lodestars is created`);
        }).catch((err) => {
          console.error('Error: Failed to write file to disk.', err);
        });
      }).catch((err) => {
        console.error('Error: Failed to read the file.', err);
      });
    }
    }
  }


  res.sendStatus(200); // Send a response back to the client
});