var EMAIL_SENT = "EMAIL_SENT";
var EDIT_MODE = "EDIT_MODE";
//Sending an email to supervisor (and student), pdf of response or edit link
function sendEmails(e) {
    //Initialize Variables
    var canSend;
    var supervisorEmail;
    var studentEmail;
    var sheet = SpreadsheetApp.getActiveSheet();
    var isEdit = false;
  var isEditing = false;
    var responseUrl = new Array();

    //Sort the responses by timestamps
    sheet.sort(1);
    var lrow = sheet.getLastRow();
    var lcol = sheet.getLastColumn();
    var lastcolumn = lcol;
    //If the last column's header is not 'Email Confirmation', make sure to make a new column at the end, to keep track of the status of
    // the response, either 'Email Sent' or 'Edit Mode'
    if (SpreadsheetApp.getActiveSheet().getRange(1, lcol).getValue() != 'Email Confirmation') {
        lastcolumn = lcol + 1;
    }
    //Set the header of the last column to Email Confirmation
    SpreadsheetApp.getActiveSheet().getRange(1, lastcolumn).setValue('Email Confirmation');
    var startRow = 2; // First row of data to process
    var dataRange = sheet.getRange(startRow, 1, lrow - 1, lastcolumn);
    //The range for the header data
    var headingsRange = sheet.getRange(1, 1, 1, lastcolumn);
    // Fetch values for each row in the header Range.
    var dataH = headingsRange.getValues();
    // Fetch values for each row in the Range.
    var data = dataRange.getValues();
    for (var i = 0; i < data.length; ++i) {
        var row = data[i];
        var emailSent = row[lastcolumn - 1];
        var rowlast = lrow - 2;
        //Response is in edit mode
        if (emailSent == EDIT_MODE && i == rowlast) {
            isEdit = true;
        }
      //Allows uers to edit (and recieve PDF) after submitting the form
       if (emailSent == EMAIL_SENT && i == rowlast) {
            isEditing = true;
        }
        // Prevents sending duplicates
        if (((emailSent != EMAIL_SENT) && (emailSent != EDIT_MODE)) || ((emailSent == EDIT_MODE) && isEdit) || isEditing) {

            //If the response is Yes, the pdf email can be sent. If no,
            //then the pdf email cannot be sent, and edit link will be sent via email, the response will be in Edit mode           
            var willSend = row[lastcolumn - 2];
            if (willSend == 'Yes') {
                canSend = true;

            } else if (willSend == 'No') {
                canSend = false;
            }

            //geting the supervisor's email
            supervisorEmail = row[1];


            if (canSend) {
                //table variables
                var tableCounters = new Array(22);
                for (var tableCount = 0; tableCount < 23; tableCount++) {
                    tableCounters[tableCount] = 0;
                }
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
                //Get the google doc file with the Ryerson header template
                var file = DriveApp.getFileById("16d4BmUZ0BrfzUNh1N4AgniIds8SolMyVxKS00Mpx1EA");
                var newId = file.getId();
                var doc = DocumentApp.openById(newId);
                var formId = (FormApp.openByUrl(SpreadsheetApp.getActiveSpreadsheet().getFormUrl())).getId();
                var form = FormApp.openById(formId);
                //Get the title of the form, so we can title the doc file
                var title = form.getTitle();
                doc.setName(title);
                var body = doc.getBody();
                body.insertParagraph(0, doc.getName())
                    .setHeading(DocumentApp.ParagraphHeading.HEADING1);
                for (var k = 0; k < dataH.length; ++k) {
                    var rowH = dataH[k];
                    var n;
                    for (n = 0; n < lastcolumn - 1; n++) {
                        //the header value
                        var contentH = rowH[n];
                        //the content value
                        var content = row[n];

                        //When the header is Please enter the student's email, obtain the student's email from the content (cell)
                        if (contentH == "Please enter the student's email.") {
                            studentEmail = content;
                        }

                        //Create the tables from the question, obtain the question by splitting the values in the header
                        //the range, data, doc, and questions and answers will be send to the createTable function
                        //Increase the counter of the question and answer array
                        var delimeter = '[';
                        var delimeter2 = ']';
                        var delVal1 = contentH.split(delimeter);
                        if (delVal1[0] == 'Nutrition Care: 3.01 Assess nutrition related risks and needs ') {
                            if (delVal1[1] != undefined) {
                                var delVal2 = delVal1[1].split(delimeter2);
                                tableQuestionCell[tableCounters[0]] = delVal2[0];
                                tableAnswerCell[tableCounters[0]] = content;
                                var tableCols = getTableCols(headingsRange, dataH, lastcolumn, delVal1[0]);
                                tableCounters[0] += 1;
                                if (tableCounters[0] == tableCols) {
                                    createTable(headingsRange, dataH, lastcolumn, data, body, tableCols, delVal1[0], tableQuestionCell, tableAnswerCell);
                                }
                            }
                        } else if (delVal1[0] == 'Nutrition Care: 3.02 Develop nutrition care plans  ') {
                            if (delVal1[1] != undefined) {
                                var delVal2 = delVal1[1].split(delimeter2);
                                tableQuestionCell2[tableCounters[1]] = delVal2[0];
                                tableAnswerCell2[tableCounters[1]] = content;
                                var tableCols = getTableCols(headingsRange, dataH, lastcolumn, delVal1[0]);
                                tableCounters[1] += 1;
                                if (tableCounters[1] == tableCols) {
                                    createTable(headingsRange, dataH, lastcolumn, data, body, tableCols, delVal1[0], tableQuestionCell2, tableAnswerCell2);
                                }
                            }
                        } else if (delVal1[0] == 'Nutrition Care: 3.03 Manage implementation of nutrition care plans  ') {
                            if (delVal1[1] != undefined) {
                                var delVal2 = delVal1[1].split(delimeter2);
                                tableQuestionCell3[tableCounters[2]] = delVal2[0];
                                tableAnswerCell3[tableCounters[2]] = content;
                                var tableCols = getTableCols(headingsRange, dataH, lastcolumn, delVal1[0]);
                                tableCounters[2] += 1;
                                if (tableCounters[2] == tableCols) {
                                    createTable(headingsRange, dataH, lastcolumn, data, body, tableCols, delVal1[0], tableQuestionCell3, tableAnswerCell3);
                                }
                            }
                        } else if (delVal1[0] == 'Nutrition Care: 3.04 Evaluate and modify nutrition care plans as appropriate ') {
                            if (delVal1[1] != undefined) {
                                var delVal2 = delVal1[1].split(delimeter2);
                                tableQuestionCell4[tableCounters[3]] = delVal2[0];
                                tableAnswerCell4[tableCounters[3]] = content;
                                var tableCols = getTableCols(headingsRange, dataH, lastcolumn, delVal1[0]);
                                tableCounters[3] += 1;
                                if (tableCounters[3] == tableCols) {
                                    createTable(headingsRange, dataH, lastcolumn, data, body, tableCols, delVal1[0], tableQuestionCell4, tableAnswerCell4);
                                }
                            }
                        } else if (delVal1[0] == 'Professional Practice: 1.01 Comply with federal and provincial/territorial requirements relevant to dietetic practice ') {
                            if (delVal1[1] != undefined) {
                                var delVal2 = delVal1[1].split(delimeter2);
                                tableQuestionCell5[tableCounters[4]] = delVal2[0];
                                tableAnswerCell5[tableCounters[4]] = content;
                                var tableCols = getTableCols(headingsRange, dataH, lastcolumn, delVal1[0]);
                                tableCounters[4] += 1;
                                if (tableCounters[4] == tableCols) {
                                    createTable(headingsRange, dataH, lastcolumn, data, body, tableCols, delVal1[0], tableQuestionCell5, tableAnswerCell5);
                                }
                            }
                        } else if (delVal1[0] == 'Professional Practice: 1.02 Comply with regulatory requirements relevant to dietetic practice ') {
                            if (delVal1[1] != undefined) {
                                var delVal2 = delVal1[1].split(delimeter2);
                                tableQuestionCell6[tableCounters[5]] = delVal2[0];
                                tableAnswerCell6[tableCounters[5]] = content;
                                var tableCols = getTableCols(headingsRange, dataH, lastcolumn, delVal1[0]);
                                tableCounters[5] += 1;
                                if (tableCounters[5] == tableCols) {
                                    createTable(headingsRange, dataH, lastcolumn, data, body, tableCols, delVal1[0], tableQuestionCell6, tableAnswerCell6);
                                }
                            }
                        } else if (delVal1[0] == 'Professional Practice: 1.03 Practice according to organizational requirements ') {
                            if (delVal1[1] != undefined) {
                                var delVal2 = delVal1[1].split(delimeter2);
                                tableQuestionCell7[tableCounters[6]] = delVal2[0];
                                tableAnswerCell7[tableCounters[6]] = content;
                                var tableCols = getTableCols(headingsRange, dataH, lastcolumn, delVal1[0]);
                                tableCounters[6] += 1;
                                if (tableCounters[6] == tableCols) {
                                    createTable(headingsRange, dataH, lastcolumn, data, body, tableCols, delVal1[0], tableQuestionCell7, tableAnswerCell7);
                                }
                            }
                        } else if (delVal1[0] == 'Professional Practice: 1.04 Practice within limits of individual level of professional knowledge and skills ') {
                            if (delVal1[1] != undefined) {
                                var delVal2 = delVal1[1].split(delimeter2);
                                tableQuestionCell8[tableCounters[7]] = delVal2[0];
                                tableAnswerCell8[tableCounters[7]] = content;
                                var tableCols = getTableCols(headingsRange, dataH, lastcolumn, delVal1[0]);
                                tableCounters[7] += 1;
                                if (tableCounters[7] == tableCols) {
                                    createTable(headingsRange, dataH, lastcolumn, data, body, tableCols, delVal1[0], tableQuestionCell8, tableAnswerCell8);
                                }
                            }
                        } else if (delVal1[0] == 'Professional Practice: 1.05 Address professional development needs ') {
                            if (delVal1[1] != undefined) {
                                var delVal2 = delVal1[1].split(delimeter2);
                                tableQuestionCell9[tableCounters[8]] = delVal2[0];
                                tableAnswerCell9[tableCounters[8]] = content;
                                var tableCols = getTableCols(headingsRange, dataH, lastcolumn, delVal1[0]);
                                tableCounters[8] += 1;
                                if (tableCounters[8] == tableCols) {
                                    createTable(headingsRange, dataH, lastcolumn, data, body, tableCols, delVal1[0], tableQuestionCell9, tableAnswerCell9);
                                }
                            }
                        } else if (delVal1[0] == 'Professional Practice: 1.06 Use a systematic approach to decision making ') {
                            if (delVal1[1] != undefined) {
                                var delVal2 = delVal1[1].split(delimeter2);
                                tableQuestionCell10[tableCounters[9]] = delVal2[0];
                                tableAnswerCell10[tableCounters[9]] = content;
                                var tableCols = getTableCols(headingsRange, dataH, lastcolumn, delVal1[0]);
                                tableCounters[9] += 1;
                                if (tableCounters[9] == tableCols) {
                                    createTable(headingsRange, dataH, lastcolumn, data, body, tableCols, delVal1[0], tableQuestionCell10, tableAnswerCell10);
                                }
                            }
                        } else if (delVal1[0] == 'Professional Practice: 1.07 Maintain a client-centred focus ') {
                            if (delVal1[1] != undefined) {
                                var delVal2 = delVal1[1].split(delimeter2);
                                tableQuestionCell11[tableCounters[10]] = delVal2[0];
                                tableAnswerCell11[tableCounters[10]] = content;
                                var tableCols = getTableCols(headingsRange, dataH, lastcolumn, delVal1[0]);
                                tableCounters[10] += 1;
                                if (tableCounters[10] == tableCols) {
                                    createTable(headingsRange, dataH, lastcolumn, data, body, tableCols, delVal1[0], tableQuestionCell11, tableAnswerCell11);
                                }
                            }
                        } else if (delVal1[0] == 'Professional Practice: 1.08 Manage time and workload effectively ') {
                            if (delVal1[1] != undefined) {
                                var delVal2 = delVal1[1].split(delimeter2);
                                tableQuestionCell12[tableCounters[11]] = delVal2[0];
                                tableAnswerCell12[tableCounters[11]] = content;
                                var tableCols = getTableCols(headingsRange, dataH, lastcolumn, delVal1[0]);
                                tableCounters[11] += 1;
                                if (tableCounters[11] == tableCols) {
                                    createTable(headingsRange, dataH, lastcolumn, data, body, tableCols, delVal1[0], tableQuestionCell12, tableAnswerCell12);
                                }
                            }
                        } else if (delVal1[0] == 'Professional Practice: 1.09 Use technologies to support practice ') {
                            if (delVal1[1] != undefined) {
                                var delVal2 = delVal1[1].split(delimeter2);
                                tableQuestionCell13[tableCounters[12]] = delVal2[0];
                                tableAnswerCell13[tableCounters[12]] = content;
                                var tableCols = getTableCols(headingsRange, dataH, lastcolumn, delVal1[0]);
                                tableCounters[12] += 1;
                                if (tableCounters[12] == tableCols) {
                                    createTable(headingsRange, dataH, lastcolumn, data, body, tableCols, delVal1[0], tableQuestionCell13, tableAnswerCell13);
                                }
                            }
                        } else if (delVal1[0] == 'Professional Practice: 1.10 Ensure appropriate and secure documentation ') {
                            if (delVal1[1] != undefined) {
                                var delVal2 = delVal1[1].split(delimeter2);
                                tableQuestionCell14[tableCounters[13]] = delVal2[0];
                                tableAnswerCell14[tableCounters[13]] = content;
                                var tableCols = getTableCols(headingsRange, dataH, lastcolumn, delVal1[0]);
                                tableCounters[13] += 1;
                                if (tableCounters[13] == tableCols) {
                                    createTable(headingsRange, dataH, lastcolumn, data, body, tableCols, delVal1[0], tableQuestionCell14, tableAnswerCell14);
                                }
                            }
                        } else if (delVal1[0] == 'Professional Practice: 1.11 Assess and enhance approaches to dietetic practice ') {
                            if (delVal1[1] != undefined) {
                                var delVal2 = delVal1[1].split(delimeter2);
                                tableQuestionCell15[tableCounters[14]] = delVal2[0];
                                tableAnswerCell15[tableCounters[14]] = content;
                                var tableCols = getTableCols(headingsRange, dataH, lastcolumn, delVal1[0]);
                                tableCounters[14] += 1;
                                if (tableCounters[14] == tableCols) {
                                    createTable(headingsRange, dataH, lastcolumn, data, body, tableCols, delVal1[0], tableQuestionCell15, tableAnswerCell15);
                                }
                            }
                        } else if (delVal1[0] == 'Professional Practice: 1.12 Contribute to advocacy efforts related to nutrition and health ') {
                            if (delVal1[1] != undefined) {
                                var delVal2 = delVal1[1].split(delimeter2);
                                tableQuestionCell16[tableCounters[15]] = delVal2[0];
                                tableAnswerCell16[tableCounters[15]] = content;
                                var tableCols = getTableCols(headingsRange, dataH, lastcolumn, delVal1[0]);
                                tableCounters[15] += 1;
                                if (tableCounters[15] == tableCols) {
                                    createTable(headingsRange, dataH, lastcolumn, data, body, tableCols, delVal1[0], tableQuestionCell16, tableAnswerCell16);
                                }
                            }
                        } else if (delVal1[0] == 'Professional Practice: 1.13 Participate in practice-based research ') {
                            if (delVal1[1] != undefined) {
                                var delVal2 = delVal1[1].split(delimeter2);
                                tableQuestionCell17[tableCounters[16]] = delVal2[0];
                                tableAnswerCell17[tableCounters[16]] = content;
                                var tableCols = getTableCols(headingsRange, dataH, lastcolumn, delVal1[0]);
                                tableCounters[16] += 1;
                                if (tableCounters[16] == tableCols) {
                                    createTable(headingsRange, dataH, lastcolumn, data, body, tableCols, delVal1[0], tableQuestionCell17, tableAnswerCell17);
                                }
                            }
                        } else if (delVal1[0] == 'Communication and Collaboration: 2.01 Select appropriate communication approaches ') {
                            if (delVal1[1] != undefined) {
                                var delVal2 = delVal1[1].split(delimeter2);
                                tableQuestionCell18[tableCounters[17]] = delVal2[0];
                                tableAnswerCell18[tableCounters[17]] = content;
                                var tableCols = getTableCols(headingsRange, dataH, lastcolumn, delVal1[0]);
                                tableCounters[17] += 1;
                                if (tableCounters[17] == tableCols) {
                                    createTable(headingsRange, dataH, lastcolumn, data, body, tableCols, delVal1[0], tableQuestionCell18, tableAnswerCell18);
                                }
                            }
                        } else if (delVal1[0] == 'Communication and Collaboration: 2.02 Use effective written communication skills ') {
                            if (delVal1[1] != undefined) {
                                var delVal2 = delVal1[1].split(delimeter2);
                                tableQuestionCell19[tableCounters[18]] = delVal2[0];
                                tableAnswerCell19[tableCounters[18]] = content;
                                var tableCols = getTableCols(headingsRange, dataH, lastcolumn, delVal1[0]);
                                tableCounters[18] += 1;
                                if (tableCounters[18] == tableCols) {
                                    createTable(headingsRange, dataH, lastcolumn, data, body, tableCols, delVal1[0], tableQuestionCell19, tableAnswerCell19);
                                }
                            }
                        } else if (delVal1[0] == 'Communication and Collaboration: 2.03 Use effective oral communication skills ') {
                            if (delVal1[1] != undefined) {
                                var delVal2 = delVal1[1].split(delimeter2);
                                tableQuestionCell20[tableCounters[19]] = delVal2[0];
                                tableAnswerCell20[tableCounters[19]] = content;
                                var tableCols = getTableCols(headingsRange, dataH, lastcolumn, delVal1[0]);
                                tableCounters[19] += 1;
                                if (tableCounters[19] == tableCols) {
                                    createTable(headingsRange, dataH, lastcolumn, data, body, tableCols, delVal1[0], tableQuestionCell20, tableAnswerCell20);
                                }
                            }
                        } else if (delVal1[0] == 'Communication and Collaboration: 2.04 Use effective interpersonal skills ') {
                            if (delVal1[1] != undefined) {
                                var delVal2 = delVal1[1].split(delimeter2);
                                tableQuestionCell21[tableCounters[20]] = delVal2[0];
                                tableAnswerCell21[tableCounters[20]] = content;
                                var tableCols = getTableCols(headingsRange, dataH, lastcolumn, delVal1[0]);
                                tableCounters[20] += 1;
                                if (tableCounters[20] == tableCols) {
                                    createTable(headingsRange, dataH, lastcolumn, data, body, tableCols, delVal1[0], tableQuestionCell21, tableAnswerCell21);
                                }
                            }
                        } else if (delVal1[0] == 'Communication and Collaboration: 2.05 Contribute to the learning of others ') {
                            if (delVal1[1] != undefined) {
                                var delVal2 = delVal1[1].split(delimeter2);
                                tableQuestionCell22[tableCounters[21]] = delVal2[0];
                                tableAnswerCell22[tableCounters[21]] = content;
                                var tableCols = getTableCols(headingsRange, dataH, lastcolumn, delVal1[0]);
                                tableCounters[21] += 1;
                                if (tableCounters[21] == tableCols) {
                                    createTable(headingsRange, dataH, lastcolumn, data, body, tableCols, delVal1[0], tableQuestionCell22, tableAnswerCell22);
                                }
                            }
                        } else if (delVal1[0] == 'Communication and Collaboration: 2.06 Contribute productively to teamwork and collaborative processes ') {
                            if (delVal1[1] != undefined) {
                                var delVal2 = delVal1[1].split(delimeter2);
                                tableQuestionCell23[tableCounters[22]] = delVal2[0];
                                tableAnswerCell23[tableCounters[22]] = content;
                                var tableCols = getTableCols(headingsRange, dataH, lastcolumn, delVal1[0]);
                                tableCounters[22] += 1;
                                if (tableCounters[22] == tableCols) {
                                    createTable(headingsRange, dataH, lastcolumn, data, body, tableCols, delVal1[0], tableQuestionCell23, tableAnswerCell23);
                                }
                            }
                        }
                        //When the data is not apart of a table, style the header and content
                        else {
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
                //Save and close the doc file
                doc.saveAndClose();
                //If the pdf email can be sent, create a pdf of the document
                //Create the email, with a subject, title, and a pdf attachment
                //Send it to the supervisor and cc the student
                //Set the last column's value to Email Sent

                var subject = "Nutrition Care";
                var pdf = DriveApp.getFileById(doc.getId()).getBlob().getAs('application/pdf').setName('Performance Evaluation Form');
                MailApp.sendEmail(supervisorEmail, subject, 'PMDip Student Performance Evaluation Form (Nutrition Care)', {
                    cc: studentEmail,
                    attachments: [pdf]
                });
                sheet.getRange(startRow + i, lastcolumn).setValue(EMAIL_SENT);

            }
            //If the pdf email cannot be sent, get the edit url
            //Create the email, with a subject, title, and the edit url
            //Send it to the supervisor
            //Set the last column to Edit Mode
            else if (canSend == false) {
                responseUrl = assignEditUrls();
                var emailTo = supervisorEmail;
                var subject = "Nutrition Care";
                var options = {}
                var theLast = lrow - 2;
                //Looks at the most recent response
                if ((i) == theLast) {
                    options.htmlBody = "PMDip Student Performance Evaluation Form" + '<br />' + '<a href=\"' + responseUrl[theLast] + '">Edit your response</a>' + ".";
                    MailApp.sendEmail(emailTo, subject, '', options);
                    sheet.getRange(startRow + i, lastcolumn).setValue(EDIT_MODE);
                }
            }
            //Open the google doc file again
            doc = DocumentApp.openById(newId);
            //Clear the contents of the file
            doc.setText('');
            doc.setName('CCS Blank Google Doc Template');
            //Save and close the doc file
            doc.saveAndClose();
            // Make sure the cell is updated right away in case the script is interrupted
            SpreadsheetApp.flush();
        }
    }
}
//Returns the edit url
function assignEditUrls() {
    //Get the Spreadsheet from the form and the data from spreadsheet
    //Get the responses from the form
    var formId = (FormApp.openByUrl(SpreadsheetApp.getActiveSpreadsheet().getFormUrl())).getId();
    var form = FormApp.openById(formId);
    var actualSheetName = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet().getName();
    var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(actualSheetName);
    var data = sheet.getDataRange().getValues();
    var responses = form.getResponses();
    var timestamps = [],
        urls = [],
        resultUrls = [];
    //Getting the Edit Response URL
    //Updates the timestamps
    for (var i = 0; i < responses.length; i++) {
        timestamps.push(responses[i].getTimestamp().setMilliseconds(0));
        urls.push(responses[i].getEditResponseUrl());
    }
    for (var j = 1; j < data.length; j++) {

        resultUrls.push([data[j][0] ? urls[timestamps.indexOf(data[j][0].setMilliseconds(0))] : '']);
    }
    //returning the edit url
    return resultUrls;
}

//Returning the column numebers for the tables
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
//Creating the table
function createTable(headingsRange, dataH, lastcolumn, data, body, tableCols, tableHeader, tableQuestionCell, tableAnswerCell) {
    body.appendParagraph("\n");
    //The style for the header
    var headerStyle = {};
    headerStyle[DocumentApp.Attribute.BACKGROUND_COLOR] = '#004C9B';
    headerStyle[DocumentApp.Attribute.BOLD] = true;
    headerStyle[DocumentApp.Attribute.HORIZONTAL_ALIGNMENT] = DocumentApp.HorizontalAlignment.CENTER;
    headerStyle[DocumentApp.Attribute.FOREGROUND_COLOR] = '#FFFFFF';
    //Style for the question cells
    var cellStyleBold = {};
    cellStyleBold[DocumentApp.Attribute.BOLD] = true;
    cellStyleBold[DocumentApp.Attribute.FOREGROUND_COLOR] = '#000000';
    //Style for the answer cells
    var cellStyle = {};
    cellStyle[DocumentApp.Attribute.BOLD] = false;
    cellStyle[DocumentApp.Attribute.FOREGROUND_COLOR] = '#000000';
    //pagraph style
    var paraStyle = {};
    paraStyle[DocumentApp.Attribute.SPACING_AFTER] = 0;
    paraStyle[DocumentApp.Attribute.LINE_SPACING] = 1;
    //apend the header table to the body in the google docs
    var table = body.appendTable();
    var tr = table.appendTableRow();
    var td = tr.appendTableCell(tableHeader);
    table.getRow(0).getCell(0).getChild(0).asParagraph().setAttributes(headerStyle)
    td.setAttributes(headerStyle);
    //Apply the para style to each paragraph in cell
    var paraInCell = td.getChild(0).asParagraph();
    paraInCell.setAttributes(paraStyle);
    //apend the content table to the body in the google docs
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
