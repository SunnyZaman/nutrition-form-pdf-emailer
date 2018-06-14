var EMAIL_SENT = "EMAIL_SENT";
var EDIT_MODE = "EDIT_MODE";


function sendEmails(e) {


    var responseNum;
    var tableCounter = 0,
        tableCounter2 = 0;
    var tableCounter3 = 0;
    var tableCounter4 = 0;
    var tableCounter5 = 0;
    var tableCounter6 = 0;
    var tableCounter7 = 0;
    var tableCounter8 = 0;
    var tableCounter9 = 0;
    var tableCounter10 = 0;
    var tableCounter11 = 0;
    var tableCounter12 = 0;
    var tableCounter13 = 0;
    var tableCounter14 = 0;
    var tableCounter15 = 0;
    var tableCounter16 = 0;
    var tableCounter17 = 0;
    var tableCounter18 = 0;
    var tableCounter19 = 0;
    var tableCounter20 = 0;
    var tableCounter21 = 0;
    var tableCounter22 = 0;
    var tableCounter23 = 0;
    var count = 0;
    var canSend;
    var makeTable;
    var supervisorEmail;
    var studentEmail;
    var sheet = SpreadsheetApp.getActiveSheet();
    sheet.sort(1);
    var lrow = sheet.getLastRow();
    var lcol = sheet.getLastColumn();
    var lastcolumn = lcol;
    if (SpreadsheetApp.getActiveSheet().getRange(1, lcol).getValue() != 'Email Confirmation') {
        lastcolumn = lcol + 1;
    }

    var isEdit = false;



    SpreadsheetApp.getActiveSheet().getRange(1, lastcolumn).setValue('Email Confirmation');
    var startRow = 2; // First row of data to process
    var dataRange = sheet.getRange(startRow, 1, lrow - 1, lastcolumn);
    var headingsRange = sheet.getRange(1, 1, 1, lastcolumn); //added
    var dataH = headingsRange.getValues(); //added
    // Fetch values for each row in the Range.
    var data = dataRange.getValues();

    var responseUrl = new Array();
    var tableQuestionCell = new Array();
    var tableAnswerCell = new Array();
    var tableQuestionCell2 = new Array();
    var tableAnswerCell2 = new Array();
    var tableQuestionCell3 = new Array();
    var tableAnswerCell3 = new Array();
    var tableQuestionCell4 = new Array();
    var tableAnswerCell4 = new Array();
    var tableQuestionCell5 = new Array();
    var tableAnswerCell5 = new Array();
    var tableQuestionCell6 = new Array();
    var tableAnswerCell6 = new Array();
    var tableQuestionCell7 = new Array();
    var tableAnswerCell7 = new Array();
    var tableQuestionCell8 = new Array();
    var tableAnswerCell8 = new Array();
    var tableQuestionCell9 = new Array();
    var tableAnswerCell9 = new Array();
    var tableQuestionCell10 = new Array();
    var tableAnswerCell10 = new Array();
    var tableQuestionCell11 = new Array();
    var tableAnswerCell11 = new Array();
    var tableQuestionCell12 = new Array();
    var tableAnswerCell12 = new Array();
    var tableQuestionCell13 = new Array();
    var tableAnswerCell13 = new Array();
    var tableQuestionCell14 = new Array();
    var tableAnswerCell14 = new Array();
    var tableQuestionCell15 = new Array();
    var tableAnswerCell15 = new Array();
    var tableQuestionCell16 = new Array();
    var tableAnswerCell16 = new Array();
    var tableQuestionCell17 = new Array();
    var tableAnswerCell17 = new Array();
    var tableQuestionCell18 = new Array();
    var tableAnswerCell18 = new Array();
    var tableQuestionCell19 = new Array();
    var tableAnswerCell19 = new Array();
    var tableQuestionCell20 = new Array();
    var tableAnswerCell20 = new Array();
    var tableQuestionCell21 = new Array();
    var tableAnswerCell21 = new Array();
    var tableQuestionCell22 = new Array();
    var tableAnswerCell22 = new Array();
    var tableQuestionCell23 = new Array();
    var tableAnswerCell23 = new Array();




    for (var i = 0; i < data.length; ++i) {
        var row = data[i];
        var emailSent = row[lastcolumn - 1];
        var rowlast = lrow - 2;
        Logger.log("i: " + i);
        Logger.log("rowlast: " + rowlast);
        if (emailSent == EDIT_MODE && i == rowlast) {
            isEdit = true;
        }
        if (((emailSent != EMAIL_SENT) && (emailSent != EDIT_MODE)) || ((emailSent == EDIT_MODE) && isEdit)) { // Prevents sending duplicates
            var file = DriveApp.getFileById("1YOQArj2DsjM779O6q1TB4aVQjeRyGtoLmvh01-fWA6g");
            var newId = file.getId();
            var doc = DocumentApp.openById(newId);
            var formId = (FormApp.openByUrl(SpreadsheetApp.getActiveSpreadsheet().getFormUrl())).getId();
            var form = FormApp.openById(formId);
            var title = form.getTitle();
            doc.setName(title);
            var body = doc.getBody();
            body.insertParagraph(0, doc.getName())
                .setHeading(DocumentApp.ParagraphHeading.HEADING1);
            for (var k = 0; k < dataH.length; ++k) {
                var rowH = dataH[k];
                var n;
                for (n = 0; n < lastcolumn - 1; n++) {
                    var contentH = rowH[n];
                    var content = row[n];

                    if (contentH == 'Email Address') {
                        supervisorEmail = content;
                    }
                    if (contentH == "Please enter the student's email.") {
                        studentEmail = content;
                    }

                    if (contentH == 'Is this form complete?') {
                        if (content == 'Yes') {
                            canSend = true;
                            makeTable = true;
                            sheet.getRange(startRow + i, lastcolumn).setValue(EMAIL_SENT);
                        } else if (content == 'No') {
                            canSend = false;
                            makeTable = false;
                        }
                    }


                    var delimeter = '[';
                    var delimeter2 = ']';
                    var delVal1 = contentH.split(delimeter);


                    if (delVal1[0] == 'Nutrition Care: 3.01 Assess nutrition related risks and needs ') {

                        if (delVal1[1] != undefined) {
                            var delVal2 = delVal1[1].split(delimeter2);

                            tableQuestionCell[tableCounter] = delVal2[0];
                            tableAnswerCell[tableCounter] = content;

                            var tableCols = getTableCols(headingsRange, dataH, lastcolumn, delVal1[0]);
                            tableCounter++;
                            if (tableCounter == tableCols) {

                                createTable(headingsRange, dataH, lastcolumn, data, body, tableCols, delVal1[0], tableQuestionCell, tableAnswerCell);
                            }

                        }


                    } else if (delVal1[0] == 'Nutrition Care: 3.02 Develop nutrition care plans  ') {
                        if (delVal1[1] != undefined) {
                            var delVal2 = delVal1[1].split(delimeter2);
                            tableQuestionCell2[tableCounter2] = delVal2[0];
                            tableAnswerCell2[tableCounter2] = content;

                            var tableCols = getTableCols(headingsRange, dataH, lastcolumn, delVal1[0]);
                            tableCounter2++;
                            if (tableCounter2 == tableCols) {
                                createTable(headingsRange, dataH, lastcolumn, data, body, tableCols, delVal1[0], tableQuestionCell2, tableAnswerCell2);
                            }
                        }
                    } else if (delVal1[0] == 'Nutrition Care: 3.03 Manage implementation of nutrition care plans  ') {
                        if (delVal1[1] != undefined) {
                            var delVal2 = delVal1[1].split(delimeter2);
                            tableQuestionCell3[tableCounter3] = delVal2[0];
                            tableAnswerCell3[tableCounter3] = content;

                            var tableCols = getTableCols(headingsRange, dataH, lastcolumn, delVal1[0]);

                            tableCounter3++;
                            if (tableCounter3 == tableCols) {
                                createTable(headingsRange, dataH, lastcolumn, data, body, tableCols, delVal1[0], tableQuestionCell3, tableAnswerCell3);
                            }
                        }
                    } else if (delVal1[0] == 'Nutrition Care: 3.04 Evaluate and modify nutrition care plans as appropriate ') {
                        if (delVal1[1] != undefined) {
                            var delVal2 = delVal1[1].split(delimeter2);
                            tableQuestionCell4[tableCounter4] = delVal2[0];
                            tableAnswerCell4[tableCounter4] = content;

                            var tableCols = getTableCols(headingsRange, dataH, lastcolumn, delVal1[0]);

                            tableCounter4++;
                            if (tableCounter4 == tableCols) {
                                createTable(headingsRange, dataH, lastcolumn, data, body, tableCols, delVal1[0], tableQuestionCell4, tableAnswerCell4);
                            }
                        }
                    } else if (delVal1[0] == 'Professional Practice: 1.01 Comply with federal and provincial/territorial requirements relevant to dietetic practice ') {
                        if (delVal1[1] != undefined) {
                            var delVal2 = delVal1[1].split(delimeter2);
                            tableQuestionCell5[tableCounter5] = delVal2[0];
                            tableAnswerCell5[tableCounter5] = content;

                            var tableCols = getTableCols(headingsRange, dataH, lastcolumn, delVal1[0]);

                            tableCounter5++;
                            if (tableCounter5 == tableCols) {
                                createTable(headingsRange, dataH, lastcolumn, data, body, tableCols, delVal1[0], tableQuestionCell5, tableAnswerCell5);
                            }
                        }
                    } else if (delVal1[0] == 'Professional Practice: 1.02 Comply with regulatory requirements relevant to dietetic practice ') {
                        if (delVal1[1] != undefined) {
                            var delVal2 = delVal1[1].split(delimeter2);
                            tableQuestionCell6[tableCounter6] = delVal2[0];
                            tableAnswerCell6[tableCounter6] = content;

                            var tableCols = getTableCols(headingsRange, dataH, lastcolumn, delVal1[0]);

                            tableCounter6++;
                            if (tableCounter6 == tableCols) {
                                createTable(headingsRange, dataH, lastcolumn, data, body, tableCols, delVal1[0], tableQuestionCell6, tableAnswerCell6);
                            }
                        }
                    } else if (delVal1[0] == 'Professional Practice: 1.03 Practice according to organizational requirements ') {
                        if (delVal1[1] != undefined) {
                            var delVal2 = delVal1[1].split(delimeter2);
                            tableQuestionCell7[tableCounter7] = delVal2[0];
                            tableAnswerCell7[tableCounter7] = content;

                            var tableCols = getTableCols(headingsRange, dataH, lastcolumn, delVal1[0]);

                            tableCounter7++;
                            if (tableCounter7 == tableCols) {
                                createTable(headingsRange, dataH, lastcolumn, data, body, tableCols, delVal1[0], tableQuestionCell7, tableAnswerCell7);
                            }
                        }
                    } else if (delVal1[0] == 'Professional Practice: 1.04 Practice within limits of individual level of professional knowledge and skills ') {
                        if (delVal1[1] != undefined) {
                            var delVal2 = delVal1[1].split(delimeter2);
                            tableQuestionCell8[tableCounter8] = delVal2[0];
                            tableAnswerCell8[tableCounter8] = content;

                            var tableCols = getTableCols(headingsRange, dataH, lastcolumn, delVal1[0]);

                            tableCounter8++;
                            if (tableCounter8 == tableCols) {
                                createTable(headingsRange, dataH, lastcolumn, data, body, tableCols, delVal1[0], tableQuestionCell8, tableAnswerCell8);
                            }
                        }
                    } else if (delVal1[0] == 'Professional Practice: 1.05 Address professional development needs ') {
                        if (delVal1[1] != undefined) {
                            var delVal2 = delVal1[1].split(delimeter2);
                            tableQuestionCell9[tableCounter9] = delVal2[0];
                            tableAnswerCell9[tableCounter9] = content;

                            var tableCols = getTableCols(headingsRange, dataH, lastcolumn, delVal1[0]);

                            tableCounter9++;
                            if (tableCounter9 == tableCols) {
                                createTable(headingsRange, dataH, lastcolumn, data, body, tableCols, delVal1[0], tableQuestionCell9, tableAnswerCell9);
                            }
                        }
                    } else if (delVal1[0] == 'Professional Practice: 1.06 Use a systematic approach to decision making ') {
                        if (delVal1[1] != undefined) {
                            var delVal2 = delVal1[1].split(delimeter2);
                            tableQuestionCell10[tableCounter10] = delVal2[0];
                            tableAnswerCell10[tableCounter10] = content;

                            var tableCols = getTableCols(headingsRange, dataH, lastcolumn, delVal1[0]);

                            tableCounter10++;
                            if (tableCounter10 == tableCols) {
                                createTable(headingsRange, dataH, lastcolumn, data, body, tableCols, delVal1[0], tableQuestionCell10, tableAnswerCell10);
                            }
                        }
                    } else if (delVal1[0] == 'Professional Practice: 1.07 Maintain a client-centred focus ') {
                        if (delVal1[1] != undefined) {
                            var delVal2 = delVal1[1].split(delimeter2);
                            tableQuestionCell11[tableCounter11] = delVal2[0];
                            tableAnswerCell11[tableCounter11] = content;

                            var tableCols = getTableCols(headingsRange, dataH, lastcolumn, delVal1[0]);

                            tableCounter11++;
                            if (tableCounter11 == tableCols) {
                                createTable(headingsRange, dataH, lastcolumn, data, body, tableCols, delVal1[0], tableQuestionCell11, tableAnswerCell11);
                            }
                        }
                    } else if (delVal1[0] == 'Professional Practice: 1.08 Manage time and workload effectively ') {
                        if (delVal1[1] != undefined) {
                            var delVal2 = delVal1[1].split(delimeter2);
                            tableQuestionCell12[tableCounter12] = delVal2[0];
                            tableAnswerCell12[tableCounter12] = content;

                            var tableCols = getTableCols(headingsRange, dataH, lastcolumn, delVal1[0]);

                            tableCounter12++;
                            if (tableCounter12 == tableCols) {
                                createTable(headingsRange, dataH, lastcolumn, data, body, tableCols, delVal1[0], tableQuestionCell12, tableAnswerCell12);
                            }
                        }
                    } else if (delVal1[0] == 'Professional Practice: 1.09 Use technologies to support practice ') {
                        if (delVal1[1] != undefined) {
                            var delVal2 = delVal1[1].split(delimeter2);
                            tableQuestionCell13[tableCounter13] = delVal2[0];
                            tableAnswerCell13[tableCounter13] = content;

                            var tableCols = getTableCols(headingsRange, dataH, lastcolumn, delVal1[0]);

                            tableCounter13++;
                            if (tableCounter13 == tableCols) {
                                createTable(headingsRange, dataH, lastcolumn, data, body, tableCols, delVal1[0], tableQuestionCell13, tableAnswerCell13);
                            }
                        }
                    } else if (delVal1[0] == 'Professional Practice: 1.10 Ensure appropriate and secure documentation ') {
                        if (delVal1[1] != undefined) {
                            var delVal2 = delVal1[1].split(delimeter2);
                            tableQuestionCell14[tableCounter14] = delVal2[0];
                            tableAnswerCell14[tableCounter14] = content;

                            var tableCols = getTableCols(headingsRange, dataH, lastcolumn, delVal1[0]);

                            tableCounter14++;
                            if (tableCounter14 == tableCols) {
                                createTable(headingsRange, dataH, lastcolumn, data, body, tableCols, delVal1[0], tableQuestionCell14, tableAnswerCell14);
                            }
                        }
                    } else if (delVal1[0] == 'Professional Practice: 1.11 Assess and enhance approaches to dietetic practice ') {
                        if (delVal1[1] != undefined) {
                            var delVal2 = delVal1[1].split(delimeter2);
                            tableQuestionCell15[tableCounter15] = delVal2[0];
                            tableAnswerCell15[tableCounter15] = content;

                            var tableCols = getTableCols(headingsRange, dataH, lastcolumn, delVal1[0]);

                            tableCounter15++;
                            if (tableCounter15 == tableCols) {
                                createTable(headingsRange, dataH, lastcolumn, data, body, tableCols, delVal1[0], tableQuestionCell15, tableAnswerCell15);
                            }
                        }
                    } else if (delVal1[0] == 'Professional Practice: 1.12 Contribute to advocacy efforts related to nutrition and health ') {
                        if (delVal1[1] != undefined) {
                            var delVal2 = delVal1[1].split(delimeter2);
                            tableQuestionCell16[tableCounter16] = delVal2[0];
                            tableAnswerCell16[tableCounter16] = content;

                            var tableCols = getTableCols(headingsRange, dataH, lastcolumn, delVal1[0]);

                            tableCounter16++;
                            if (tableCounter16 == tableCols) {
                                createTable(headingsRange, dataH, lastcolumn, data, body, tableCols, delVal1[0], tableQuestionCell16, tableAnswerCell16);
                            }
                        }
                    } else if (delVal1[0] == 'Professional Practice: 1.13 Participate in practice-based research ') {
                        if (delVal1[1] != undefined) {
                            var delVal2 = delVal1[1].split(delimeter2);
                            tableQuestionCell17[tableCounter17] = delVal2[0];
                            tableAnswerCell17[tableCounter17] = content;

                            var tableCols = getTableCols(headingsRange, dataH, lastcolumn, delVal1[0]);

                            tableCounter17++;
                            if (tableCounter17 == tableCols) {
                                createTable(headingsRange, dataH, lastcolumn, data, body, tableCols, delVal1[0], tableQuestionCell17, tableAnswerCell17);
                            }
                        }
                    } else if (delVal1[0] == 'Communication and Collaboration: 2.01 Select appropriate communication approaches ') {
                        if (delVal1[1] != undefined) {
                            var delVal2 = delVal1[1].split(delimeter2);
                            tableQuestionCell18[tableCounter18] = delVal2[0];
                            tableAnswerCell18[tableCounter18] = content;

                            var tableCols = getTableCols(headingsRange, dataH, lastcolumn, delVal1[0]);

                            tableCounter18++;
                            if (tableCounter18 == tableCols) {
                                createTable(headingsRange, dataH, lastcolumn, data, body, tableCols, delVal1[0], tableQuestionCell18, tableAnswerCell18);
                            }
                        }
                    } else if (delVal1[0] == 'Communication and Collaboration: 2.02 Use effective written communication skills ') {
                        if (delVal1[1] != undefined) {
                            var delVal2 = delVal1[1].split(delimeter2);
                            tableQuestionCell19[tableCounter19] = delVal2[0];
                            tableAnswerCell19[tableCounter19] = content;

                            var tableCols = getTableCols(headingsRange, dataH, lastcolumn, delVal1[0]);

                            tableCounter19++;
                            if (tableCounter19 == tableCols) {
                                createTable(headingsRange, dataH, lastcolumn, data, body, tableCols, delVal1[0], tableQuestionCell19, tableAnswerCell19);
                            }
                        }
                    } else if (delVal1[0] == 'Communication and Collaboration: 2.03 Use effective oral communication skills ') {
                        if (delVal1[1] != undefined) {
                            var delVal2 = delVal1[1].split(delimeter2);
                            tableQuestionCell20[tableCounter20] = delVal2[0];
                            tableAnswerCell20[tableCounter20] = content;

                            var tableCols = getTableCols(headingsRange, dataH, lastcolumn, delVal1[0]);

                            tableCounter20++;
                            if (tableCounter20 == tableCols) {
                                createTable(headingsRange, dataH, lastcolumn, data, body, tableCols, delVal1[0], tableQuestionCell20, tableAnswerCell20);
                            }
                        }
                    } else if (delVal1[0] == 'Communication and Collaboration: 2.04 Use effective interpersonal skills ') {
                        if (delVal1[1] != undefined) {
                            var delVal2 = delVal1[1].split(delimeter2);
                            tableQuestionCell21[tableCounter21] = delVal2[0];
                            tableAnswerCell21[tableCounter21] = content;

                            var tableCols = getTableCols(headingsRange, dataH, lastcolumn, delVal1[0]);

                            tableCounter21++;
                            if (tableCounter21 == tableCols) {
                                createTable(headingsRange, dataH, lastcolumn, data, body, tableCols, delVal1[0], tableQuestionCell21, tableAnswerCell21);
                            }
                        }
                    } else if (delVal1[0] == 'Communication and Collaboration: 2.05 Contribute to the learning of others ') {
                        if (delVal1[1] != undefined) {
                            var delVal2 = delVal1[1].split(delimeter2);
                            tableQuestionCell22[tableCounter22] = delVal2[0];
                            tableAnswerCell22[tableCounter22] = content;

                            var tableCols = getTableCols(headingsRange, dataH, lastcolumn, delVal1[0]);

                            tableCounter22++;
                            if (tableCounter22 == tableCols) {
                                createTable(headingsRange, dataH, lastcolumn, data, body, tableCols, delVal1[0], tableQuestionCell22, tableAnswerCell22);
                            }
                        }
                    } else if (delVal1[0] == 'Communication and Collaboration: 2.06 Contribute productively to teamwork and collaborative processes ') {
                        if (delVal1[1] != undefined) {
                            var delVal2 = delVal1[1].split(delimeter2);
                            tableQuestionCell23[tableCounter23] = delVal2[0];
                            tableAnswerCell23[tableCounter23] = content;

                            var tableCols = getTableCols(headingsRange, dataH, lastcolumn, delVal1[0]);

                            tableCounter23++;
                            if (tableCounter23 == tableCols) {
                                createTable(headingsRange, dataH, lastcolumn, data, body, tableCols, delVal1[0], tableQuestionCell23, tableAnswerCell23);
                            }
                        }
                    } else {
                        var style = {};
                        style[DocumentApp.Attribute.HORIZONTAL_ALIGNMENT] =
                            DocumentApp.HorizontalAlignment.LEFT;
                        style[DocumentApp.Attribute.FONT_FAMILY] = 'Calibri';
                        style[DocumentApp.Attribute.FONT_SIZE] = 18;
                        style[DocumentApp.Attribute.BOLD] = true;
                        style[DocumentApp.Attribute.FOREGROUND_COLOR] = '#104599';
                        var style2 = {};
                        style2[DocumentApp.Attribute.HORIZONTAL_ALIGNMENT] =
                            DocumentApp.HorizontalAlignment.LEFT;
                        style2[DocumentApp.Attribute.FONT_FAMILY] = 'Calibri';
                        style2[DocumentApp.Attribute.FONT_SIZE] = 12;
                        style2[DocumentApp.Attribute.BOLD] = false;
                        style2[DocumentApp.Attribute.FOREGROUND_COLOR] = '#000000';
                        var headcont = body.appendParagraph(contentH);
                        headcont.setAttributes(style);
                        var cont = body.appendParagraph(content);
                        cont.setAttributes(style2);


                    }

                }


            }
            doc.saveAndClose();


            if (canSend) {
                count = 0;
                var subject = "Nutrition Care";
                var pdf = DriveApp.getFileById(doc.getId()).getBlob().getAs('application/pdf').setName('Performance Evaluation Form');
                MailApp.sendEmail(supervisorEmail, subject, 'PMDip Student Performance Evaluation Form (Nutrition Care)', {
                    cc: studentEmail,
                    attachments: [pdf]
                });
                sheet.getRange(startRow + i, lastcolumn).setValue(EMAIL_SENT);
                var testLog = sheet.getRange(startRow + i, lastcolumn).getValue();

            } else if (canSend == false) {
                responseUrl = assignEditUrls();
                var emailTo = supervisorEmail;
                var subject = "Nutrition Care";
                var options = {}
                var theLast = lrow - 2;

                if ((i) == theLast) {
                    options.htmlBody = "PMDip Student Performance Evaluation Form" + '<br />' + '<a href=\"' + responseUrl[theLast] + '">Edit your response</a>' + ".";
                    MailApp.sendEmail(emailTo, subject, '', options);
                    sheet.getRange(startRow + i, lastcolumn).setValue(EDIT_MODE);
                }
            }

            doc = DocumentApp.openById(newId);
            doc.setText('');
            doc.setName('CCS Blank Google Doc Template');
            doc.saveAndClose();
            // Make sure the cell is updated right away in case the script is interrupted
            SpreadsheetApp.flush();
        } 

    }
}



function assignEditUrls() {
    var formId = (FormApp.openByUrl(SpreadsheetApp.getActiveSpreadsheet().getFormUrl())).getId();
    var form = FormApp.openById(formId);
    //enter form ID here

    var actualSheetName = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet().getName();
    var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(actualSheetName);

    //Change the sheet name as appropriate
    var data = sheet.getDataRange().getValues();
    var urlCol = 6; // column number where URL's should be populated; A = 1, B = 2 etc
    var responses = form.getResponses();
    var timestamps = [],
        urls = [],
        resultUrls = [];

    for (var i = 0; i < responses.length; i++) {
        timestamps.push(responses[i].getTimestamp().setMilliseconds(0));
        urls.push(responses[i].getEditResponseUrl());
    }
    for (var j = 1; j < data.length; j++) {

        resultUrls.push([data[j][0] ? urls[timestamps.indexOf(data[j][0].setMilliseconds(0))] : '']);
    }

    return resultUrls;
}

function getTableCols(headingsRange, dataH, lastcolumn, theHeader) {
    var columnCounter = 0;
    for (var k = 0; k < dataH.length; ++k) {
        var rowH = dataH[k];
        for (var n = 0; n < lastcolumn - 1; n++) {
            var contentH = rowH[n]
            var delimeter = '[';
            var delVal1 = contentH.split(delimeter);
            if (delVal1[0] == theHeader) {
                columnCounter++;
            }
        }
    }
    return columnCounter;
}

function createTable(headingsRange, dataH, lastcolumn, data, body, tableCols, tableHeader, tableQuestionCell, tableAnswerCell) {

    body.appendParagraph("\n");
    //style
    var headerStyle = {};
    headerStyle[DocumentApp.Attribute.BACKGROUND_COLOR] = '#004C9B';
    headerStyle[DocumentApp.Attribute.BOLD] = true;
    headerStyle[DocumentApp.Attribute.HORIZONTAL_ALIGNMENT] = DocumentApp.HorizontalAlignment.CENTER;
    headerStyle[DocumentApp.Attribute.FOREGROUND_COLOR] = '#FFFFFF';

    //Style for the cells other than header row
    var cellStyleBold = {};
    cellStyleBold[DocumentApp.Attribute.BOLD] = true;
    cellStyleBold[DocumentApp.Attribute.FOREGROUND_COLOR] = '#000000';

    //Style for the cells other than header row
    var cellStyle = {};
    cellStyle[DocumentApp.Attribute.BOLD] = false;
    cellStyle[DocumentApp.Attribute.FOREGROUND_COLOR] = '#000000';

    var paraStyle = {};
    paraStyle[DocumentApp.Attribute.SPACING_AFTER] = 0;
    paraStyle[DocumentApp.Attribute.LINE_SPACING] = 1;

    var table = body.appendTable();

    var tr = table.appendTableRow();

    var td = tr.appendTableCell(tableHeader);
    table.getRow(0).getCell(0).getChild(0).asParagraph().setAttributes(headerStyle)
    td.setAttributes(headerStyle);


    //Apply the para style to each paragraph in cell
    var paraInCell = td.getChild(0).asParagraph();
    paraInCell.setAttributes(paraStyle);



    var table2 = body.appendTable();
    for (var l = 0; l < tableCols; l++) {
        var tr = table2.appendTableRow();
        for (var m = 0; m < 1; m++) {
            var td = tr.appendTableCell(tableQuestionCell[l]);
            var td2 = tr.appendTableCell(tableAnswerCell[l]);
            td.setAttributes(cellStyle);
            td.setAttributes(cellStyle);
            var paraInCell = td.getChild(0).asParagraph();
            paraInCell.setAttributes(paraStyle);
        }

    }
}
