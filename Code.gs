function doGet(e) {
    var template = HtmlService.createTemplateFromFile('Page');
    return template.evaluate()
        .setTitle('Create Random Assignments 2.0')
        .setSandboxMode(HtmlService.SandboxMode.IFRAME);
}

function getOAuthToken() {
    DriveApp.getRootFolder();
    return ScriptApp.getOAuthToken();
}

function CheckData(data) {
    return interfaceClass.submitWindow(data);
}


function onlyUnique(value, index, self) {
    return self.indexOf(value) === index;
}

function isInt(n) {
    return parseFloat(n) == parseInt(n, 10) && !isNaN(n);
}

function isFloat(n) {
    return !isNaN(n);
}

function validateEmail(email) {
    var re = /^(([^<>()[\]\\.,;:\s@\"]+(\.[^<>()[\]\\.,;:\s@\"]+)*)|(\".+\"))@((\[[0-9]{1,3}\.[0-9]{1,3}\.[0-9]{1,3}\.[0-9]{1,3}\])|(([a-zA-Z\-0-9]+\.)+[a-zA-Z]{2,}))$/;
    return re.test(email);
}

function getProgress() {
    var userProperties = PropertiesService.getUserProperties();
    return " (" + userProperties.getProperty('studentnum') + "/" + userProperties.getProperty('studenttotal') + ")";
}


function GoogleBugWorkAround(title, folder) {
    var doc = DocumentApp.create(title);
    var docid = doc.getId();
    doc.saveAndClose();
    var myfile = DriveApp.getFileById(docid);
    folder.addFile(myfile);
    return myfile;
}


function shuffle(array) {
    var counter = array.length,
        temp, index;

    // While there are elements in the array
    while (counter--) {
        // Pick a random index
        index = (Math.random() * counter) | 0;

        // And swap the last element with it
        temp = array[counter];
        array[counter] = array[index];
        array[index] = temp;
    }

    return array;
}



var interfaceClass = {};

interfaceClass.assignmentName = "";
interfaceClass.assignmentFile = "";
interfaceClass.studentFile = "";
interfaceClass.numberOfQuestions = 1;
interfaceClass.questionfileList = [];
interfaceClass.studentfileList = [];
interfaceClass.notifyStudents = "";
interfaceClass.allowLink = "";
interfaceClass.individualFiles = "";
interfaceClass.iak = "";
interfaceClass.readonly = "";
interfaceClass.sendfile = "";
interfaceClass.studentnum = 0;
interfaceClass.studenttotal = 0;
interfaceClass.quota = 1500;
interfaceClass.headerbody = "";



interfaceClass.submitWindow = function(data) {
    interfaceClass.studentFile = "";
    interfaceClass.assignmentFile = "";
    interfaceClass.headerbody = "";
    
    interfaceClass.quota = MailApp.getRemainingDailyQuota();
    if (data.students.trim().length > 0) {
        var fileCheck = data.students.trim();
        var check = walkQuestionSheet.checkFileType(fileCheck, MimeType.GOOGLE_SHEETS);
        if (check === true) {
            interfaceClass.studentFile = fileCheck;
        }
    }
    if (data.questions.trim().length > 0) {
        var fileCheck = data.questions.trim();
        var check = walkQuestionSheet.checkFileType(fileCheck, MimeType.GOOGLE_SHEETS);
        if (check === true) {
            interfaceClass.assignmentFile = fileCheck;
        }
    }

    var email = (data.delivery === 'email') ? true : false;
    var ind = (data.delivery === 'ind') ? true : false;
    var single = (data.delivery === 'single') ? true : false;
    var assignmentName = data.assignment.trim();
    var numberofquestions = data.numq.trim();
    var attachment = (data.attachment === "on") ? true : false;
    var link = (data.link === "on") ? true : false;
    var readonly = (data.readonly === "on") ? true : false;
    var answerkey = (data.answerkeys === "on") ? true : false;

    var subjectLine = data.subjectline.trim();
    var headerbody = data.headerbody.trim();
    if (subjectLine.length > 0) {
        subjectLine = subjectLine;
    } else {
        subjectLine = false;
    }
    
    if (headerbody.length > 0){
        headerbody = headerbody;
    } else {
        headerbody = false;
    }
    interfaceClass.headerbody = headerbody;

    var start = data.lStart;
	var end = data.lEnd;
  
    var rangecheck = false;
    if (isInt(start) && isInt(end)) {
      var start = parseInt(start, 10);
      var end = parseInt(end, 10);
      if (start > 0 && end >= start) {
            var rangecheck = true;
      }
    }
    if (assignmentName.length > 0 && interfaceClass.assignmentFile.length > 0 && interfaceClass.studentFile.length > 0 && isInt(numberofquestions)) {
        if (isInt(numberofquestions)) {
            interfaceClass.numberOfQuestions = parseInt(numberofquestions, 10);
            walkQuestionSheet.numberofquestionsdefault = interfaceClass.numberOfQuestions;
        }
        interfaceClass.assignmentName = assignmentName;
        if (email === true || ind === true) {
            interfaceClass.individualFiles = true;

            if (link === true) {
                interfaceClass.allowLink = true;
            } else {
                interfaceClass.allowLink = false;
            }
            if (email === true) {
                interfaceClass.notifyStudents = true;
            } else {
                interfaceClass.notifyStudents = false;
            }
            if (answerkey === true) {
                interfaceClass.iak = true;
            } else {
                interfaceClass.iak = false;
            }
            if (attachment === true) {
                interfaceClass.sendfile = true;
            } else {
                interfaceClass.sendfile = false;
            }
            if (readonly === true) {
                interfaceClass.readonly = true;
            } else {
                interfaceClass.readonly = false;
            }


        } else {
            interfaceClass.individualFiles = false;
            interfaceClass.allowLink = false;
            interfaceClass.notifyStudents = false;
            interfaceClass.readonly = false;
            interfaceClass.iak = false;
        }
        result = interfaceClass.runAssignment(rangecheck, start, end, subjectLine);
        if (result === true) {
            return result;
        } else if (typeof result == 'string') {
            return result;
        }
    }
    return "Something is wrong with the parameter specified.";
};

interfaceClass.runAssignment = function(range, start, end, subjectLine) {
    if (interfaceClass.assignmentName.length > 0 && interfaceClass.studentFile.length > 0 && interfaceClass.assignmentFile.length > 0) {

        var studentSpreadSheet = walkQuestionSheet.getStudentHeader(interfaceClass.studentFile);
        var questionSpreadSheet = walkQuestionSheet.getHeader(interfaceClass.assignmentFile);
        if (studentSpreadSheet === false) {
            return "Something is wrong with the student spreadsheet specified.";
        }
        if (questionSpreadSheet === false) {
            return "Something is wrong with the question spreadsheet specified.";
        }
        var initVal = 0;
        var endval = walkQuestionSheet.studentNames.length;
        if (range === true) {
            if (start > 0 && start <= walkQuestionSheet.studentNames.length) {
                var initVal = start - 1;
            }
            if (end > 0 && end <= walkQuestionSheet.studentNames.length) {
                if (end >= (initVal + 1)) {
                    endval = end;
                } else {
                    endval = initVal + 1;
                }
            }
        }

        var folder = newDocumentClass.createFolder(interfaceClass.assignmentName);
        //		var answerkey = folder.createFile('Answer Key', '', MimeType.GOOGLE_DOCS);
        var answerkey = GoogleBugWorkAround('Answer Key', folder);
        answerkey.setStarred(true);
        if (interfaceClass.individualFiles !== true) {
            //			var questionsheet = folder.createFile('Question Sheet', '', MimeType.GOOGLE_DOCS);
            var questionsheet = GoogleBugWorkAround('Question Sheet', folder);
        }
        var userProperties = PropertiesService.getUserProperties();
        interfaceClass.studenttotal = endval;
        interfaceClass.studentnum = initVal + 1;
        numstudentstosend = endval-initVal;
        if (numstudentstosend > interfaceClass.quota && interfaceClass.notifyStudents === true && interfaceClass.individualFiles === true){
          return 'You are attempting to send ' + Math.round(numstudentstosend) + ' emails while your daily remaining Gmail sending quota is currently at ' + Math.round(interfaceClass.quota) + '. You can set a range of students you would like to send emails to under the advanced settings section.';
        }
      
        userProperties.setProperty('studenttotal', interfaceClass.studenttotal.toString());
        userProperties.setProperty('studentnum', interfaceClass.studentnum.toString());
        for (var i = initVal; i < endval; i++) {
            interfaceClass.studentnum = i + 1;

            if (interfaceClass.studentnum % 10 === 0) {
                userProperties.setProperty('studentnum', interfaceClass.studentnum.toString());
            }
            var sName = walkQuestionSheet.studentNames[i];
            if (walkQuestionSheet.studentEmailCol > -1) {
                var sEmail = walkQuestionSheet.studentEmails[i];
            } else {
                var sEmail = "";
            }
            if (i > 0) {
                var newPage = true;
            } else {
                var newPage = false;
            }
            var iakdocument = false;
            if (interfaceClass.individualFiles === true) {
                var questionsheet = newDocumentClass.createFile(folder, interfaceClass.assignmentName, sName, sEmail);
                if (interfaceClass.iak === true) {
                    var iakdocument = newDocumentClass.createIAKFile(folder, interfaceClass.assignmentName, sName, sEmail);
                }
            }
            interfaceClass.createStudent(questionSpreadSheet, answerkey, questionsheet, sName, sEmail, newPage, iakdocument);

            if (interfaceClass.individualFiles === true && interfaceClass.notifyStudents === true && validateEmail(sEmail)) {
                newEmailClass.sendEmail(questionsheet, interfaceClass.assignmentName, "The following document has been sent to you:\n\n", sName, sEmail, subjectLine);
            }
        }
        userProperties.setProperty('studenttotal', '');
        userProperties.setProperty('studentnum', '');
        return true;
    }
    return false;
};


interfaceClass.createStudent = function(qsheet, answersheet, questionsheet, name, email, newPage, iak) {
    if (typeof iak === "undefined") {
        iak = false;
    }
    var allQuestions = [];
    if (walkQuestionSheet.chapterCol > -1) {
        for (var i = 0; i < walkQuestionSheet.chapters.length; i++) {
            var availableQuestions = walkQuestionSheet.chapterQuestions[i];
            var questionCount = walkQuestionSheet.numberofquestions[i];
            var reqQuestions = walkQuestionSheet.chapterRQuestions[i];
            var randomQuestions = newDocumentClass.getRandomQuestions(availableQuestions, questionCount);
            var chapterQuestions = randomQuestions.concat(reqQuestions);
            chapterQuestions = shuffle(chapterQuestions);
            var allQuestions = allQuestions.concat(chapterQuestions);
        }
    } else {
        var availableQuestions = walkQuestionSheet.chapterQuestions;
        var questionCount = walkQuestionSheet.numberofquestions;
        var reqQuestions = walkQuestionSheet.chapterRQuestions;
        var randomQuestions = newDocumentClass.getRandomQuestions(availableQuestions, questionCount);
        var chapterQuestions = randomQuestions.concat(reqQuestions);
        var allQuestions = allQuestions.concat(chapterQuestions);
        allQuestions = shuffle(allQuestions);
    }
    var displayname = name;
    if (email.length > 0) {
        var displayname = displayname + ' - ' + email;
    }
    var shell = newDocumentClass.createShell(answersheet, interfaceClass.assignmentName, displayname, newPage);

    if (interfaceClass.individualFiles === true) {
        var qshell = newDocumentClass.createShell(questionsheet, interfaceClass.assignmentName, displayname, false);
    } else {
        var qshell = newDocumentClass.createShell(questionsheet, interfaceClass.assignmentName, displayname, newPage);
    }

    //	allQuestions = shuffle(allQuestions);


    var questions = [];
    var answers = [];
    for (var i = 0; i < allQuestions.length; i++) {
        indexVal = allQuestions[i];
        questions.push(walkQuestionSheet.questionText[indexVal]);
        if (walkQuestionSheet.answerCol > -1) {
            answers.push(walkQuestionSheet.answerText[indexVal]);
        } else {
            answers.push("");
        }
    }

    newDocumentClass.createAnswers(shell, questions, answers);
    newDocumentClass.createQuestions(qshell, questions);
    if (interfaceClass.individualFiles === true) {
        qshell.saveAndClose();
    }
    if (iak !== false && interfaceClass.iak === true) {
        var iakshell = newDocumentClass.createShell(iak, interfaceClass.assignmentName, displayname, false);
        newDocumentClass.createAnswers(iakshell, questions, answers);
        iakshell.saveAndClose();
    }

    return true;
};

var newDocumentClass = {};

newDocumentClass.createFolder = function(folderName) {
    var folder = DriveApp.createFolder(folderName);
    return folder;
};

newDocumentClass.getRandomQuestions = function(itemArray, numItems) {
    var clone = itemArray.slice(0);
    var arrayReturn = [];
    clone = shuffle(clone);
    while (arrayReturn.length < numItems && clone.length > 0) {
        arrayReturn.push(clone.pop());
    }
    return arrayReturn;
};


newDocumentClass.createIAKFile = function(folder, assignment, studentName, studentEmail) {
    var name = assignment + ' - ' + studentName + ' - Key';
    if (studentEmail.length > 0) {
        var name = name + ' - ' + studentEmail;
    }
    //	var file = folder.createFile(name, '', MimeType.GOOGLE_DOCS);
    var file = GoogleBugWorkAround(name, folder);
    return file;
};

newDocumentClass.createFile = function(folder, assignment, studentName, studentEmail) {
    var name = assignment + ' - ' + studentName;
    if (studentEmail.length > 0) {
        var name = name + ' - ' + studentEmail;
    }
    var file = GoogleBugWorkAround(name, folder);
    if (interfaceClass.allowLink != true && validateEmail(studentEmail)) {
        if (interfaceClass.readonly == true) {  
          try {
              file.addViewer(String(studentEmail));
          } catch (error) {
              console.error(error);
          }
        } else {
          try {
              file.addEditor(String(studentEmail));
          } catch (error) {
              console.error(error);
          }
        }
    } else if (interfaceClass.allowLink == true) {
        if (interfaceClass.readonly) {
            file.setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.VIEW);
        } else {
            file.setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.EDIT);
        }
    }
    return file;
};

newDocumentClass.createShell = function(file, assignment, student, newPage) {

    if (typeof newPage === "undefined") {
        newPage = false;
    }

    var fileid = file.getId();
    var doc = DocumentApp.openById(fileid);

    var title = assignment;
    var body = doc.getBody();

    if (newPage == true) {
        body.appendPageBreak()
    }
    var bheader = interfaceClass.headerbody;
    var reportTitle = body.appendParagraph(title);
    reportTitle.setFontFamily(DocumentApp.FontFamily.ARIAL);
    reportTitle.setFontSize(24);
    reportTitle.setForegroundColor('#4a86e8');
    reportTitle.setAlignment(DocumentApp.HorizontalAlignment.CENTER);
    reportTitle.setBold(false);
    reportTitle.setItalic(false);

    var overview = body.appendParagraph(student);
    overview.setFontSize(14);
    overview.setSpacingBefore(14);
    overview.setBold(true);
    overview.setItalic(false);

    overview.setForegroundColor('#000000');
  
  
    if (bheader != false){
       var bodyheader = body.appendParagraph(bheader);
       bodyheader.setFontFamily(DocumentApp.FontFamily.ARIAL);
       bodyheader.setFontSize(12);
       bodyheader.setSpacingBefore(6);
       bodyheader.setForegroundColor('#000000');
       bodyheader.setBold(false);
       bodyheader.setItalic(true);
    }

    /*
    var footer = doc.addFooter();

    var divider = footer.appendHorizontalRule();

    var footerText = footer.appendParagraph('Confidential and proprietary');
    footerText.setFontSize(9);
    footerText.setForegroundColor('#4a86e8');
    footerText.setAlignment(DocumentApp.HorizontalAlignment.RIGHT);
    */

    var questions = body.appendParagraph('Questions');
    questions.setFontSize(14);
    questions.setSpacingBefore(14);
    questions.setBold(true);
    questions.setItalic(false);
    questions.setForegroundColor('#000000');
    return doc;
};

newDocumentClass.createQuestions = function(doc, questions) {
    var body = doc.getBody();
    for (var i = 0; i < questions.length; i++) {
        content = String(i + 1) + ") " + questions[i];

        var question = body.appendParagraph(content);
        question.setFontFamily(DocumentApp.FontFamily.ARIAL);
        question.setFontSize(12);
        question.setSpacingBefore(6);
        question.setForegroundColor('#000000');
        question.setBold(false);
        question.setItalic(false);

        var question = body.appendParagraph("");
        question.setFontFamily(DocumentApp.FontFamily.ARIAL);
        question.setFontSize(12);
        question.setSpacingBefore(6);

        var divider = body.appendHorizontalRule();
    }
    return doc;
};

newDocumentClass.createAnswers = function(doc, questions, answers) {
    var body = doc.getBody();
    for (var i = 0; i < questions.length; i++) {
        content = String(i + 1) + ") " + questions[i];
        answerc = answers[i];

        var question = body.appendParagraph(content);
        question.setFontFamily(DocumentApp.FontFamily.ARIAL);
        question.setFontSize(12);
        question.setSpacingBefore(6);
        question.setBold(false);
        question.setItalic(false);
        question.setForegroundColor('#000000');

        var answer = body.appendParagraph(answerc);
        answer.setFontFamily(DocumentApp.FontFamily.ARIAL);
        answer.setFontSize(12);
        answer.setSpacingBefore(6);
        answer.setItalic(true);
        answer.setForegroundColor('#000000');
        answer.setBold(false);

        var divider = body.appendHorizontalRule();
    }
    return doc;
};


walkQuestionSheet = {};

walkQuestionSheet.questionCol = -1;
walkQuestionSheet.requiredCol = -1;
walkQuestionSheet.chapterCol = -1;
walkQuestionSheet.questionCountCol = -1;
walkQuestionSheet.chapters = [];
walkQuestionSheet.chapterRQuestions = [];
walkQuestionSheet.chapterQuestions = [];
walkQuestionSheet.numberofquestions = [];
walkQuestionSheet.numberofquestionsdefault = 1;
walkQuestionSheet.answerCol = -1;

walkQuestionSheet.studentNameCol = -1;
walkQuestionSheet.studentEmailCol = -1;
walkQuestionSheet.studentNames = [];
walkQuestionSheet.studentEmails = [];

walkQuestionSheet.questionText = [];
walkQuestionSheet.answerText = [];



walkQuestionSheet.checkFileType = function(myfile, reqType) {
    var file = DriveApp.getFileById(myfile);
    if (file.getMimeType() == reqType) {
        return true;
    } else {
        return false;
    }
};


walkQuestionSheet.getFolders = function() {
    var folders = DriveApp.getFolders();
    var fList = [];
    while (folders.hasNext()) {
        var folder = folders.next();
        var fItem = {};
        fItem.id = folder.getId();
        fItem.name = folder.getName();
        fList.push(fItem);
    }
    return folders;
};


walkQuestionSheet.listSpreadSheets = function(folder) {
    if (typeof folder === "undefined") {
        var files = DriveApp.getFilesByType(MimeType.GOOGLE_SHEETS);
    } else {
        var files = folder.getFilesByType(MimeType.GOOGLE_SHEETS);
    }
    var fList = [];

    while (files.hasNext()) {
        var doc = files.next();
        var fItem = {};
        fItem.id = doc.getId();
        fItem.name = doc.getName();
        fList.push(fItem);
    }

    return fList;
};


walkQuestionSheet.getStudentHeader = function(file) {
    var spreadsheet = SpreadsheetApp.openById(file);
    var sheets = spreadsheet.getSheets();
    if (sheets.length > 0) {
        var firstSheet = sheets[0];
        var data = firstSheet.getDataRange()
            .getValues();
        if (data.length > 0) {
            var firstRow = data[0];
            for (var i = 0; i < firstRow.length; i++) {
                if (firstRow[i].toLowerCase() == 'name') {
                    walkQuestionSheet.studentNameCol = i;
                } else if (firstRow[i].toLowerCase() == 'email' || firstRow[i].toLowerCase() == 'e-mail' || firstRow[i].toLowerCase() == 'e-mail address' || firstRow[i].toLowerCase() == 'email address') {
                    walkQuestionSheet.studentEmailCol = i;
                }
            }
        }
        if (walkQuestionSheet.studentNameCol > -1) {
            for (var i = 1; i < data.length; i++) {
                walkQuestionSheet.studentNames.push(data[i][walkQuestionSheet.studentNameCol]);
                if (walkQuestionSheet.studentEmailCol > -1) {
                    walkQuestionSheet.studentEmails.push(data[i][walkQuestionSheet.studentEmailCol]);
                }
            }
        }
        if (walkQuestionSheet.studentNameCol == -1) {
            return false;
        }
        return firstSheet;
    }
    return false;
};

walkQuestionSheet.getValue = function(sheetObj, row, col) {
    var firstSheet = sheetObj;
    var data = firstSheet.getDataRange()
        .getValues();
    var r = parseInt(row);
    var c = parseInt(col);
    return data[r][c];
};


walkQuestionSheet.getHeader = function(file) {
    var spreadsheet = SpreadsheetApp.openById(file);
    var sheets = spreadsheet.getSheets();
    if (sheets.length > 0) {
        var firstSheet = sheets[0];
        var data = firstSheet.getDataRange()
            .getValues();

        if (data.length > 0) {
            var firstRow = data[0];
            for (var i = 0; i < firstRow.length; i++) {
                if (firstRow[i].toLowerCase() == 'question' || firstRow[i].toLowerCase() == 'questions') {
                    walkQuestionSheet.questionCol = i;
                } else if (firstRow[i].toLowerCase() == 'chapter' || firstRow[i].toLowerCase() == 'chapters' || firstRow[i].toLowerCase() == 'section' || firstRow[i].toLowerCase() == 'sections') {
                    walkQuestionSheet.chapterCol = i;
                } else if (firstRow[i].toLowerCase() == 'required') {
                    walkQuestionSheet.requiredCol = i;
                } else if (firstRow[i].toLowerCase() == 'number' || firstRow[i].toLowerCase() == 'questions per chapter' || firstRow[i].toLowerCase() == 'question per chapter') {
                    walkQuestionSheet.questionCountCol = i;
                } else if (firstRow[i].toLowerCase() == 'answer') {
                    walkQuestionSheet.answerCol = i;
                }
            }
        }
        if (walkQuestionSheet.questionCol > -1) {
            walkQuestionSheet.questionText.push("");
            walkQuestionSheet.answerText.push("");
            for (var i = 1; i < data.length; i++) {
                walkQuestionSheet.questionText.push(data[i][walkQuestionSheet.questionCol]);
                if (walkQuestionSheet.answerCol > -1) {
                    walkQuestionSheet.answerText.push(data[i][walkQuestionSheet.answerCol]);
                } else {
                    walkQuestionSheet.answerText.push("");
                }
            }
        }
        if (walkQuestionSheet.chapterCol > -1) {
            for (var i = 1; i < data.length; i++) {
                if (walkQuestionSheet.chapters.indexOf(data[i][walkQuestionSheet.chapterCol]) > -1) {
                    cqi = walkQuestionSheet.chapters.indexOf(data[i][walkQuestionSheet.chapterCol]);
                    if (walkQuestionSheet.requiredCol > -1) {
                        if (data[i][walkQuestionSheet.requiredCol].toLowerCase() == 'true' || data[i][walkQuestionSheet.requiredCol].toLowerCase() == 't' || data[i][walkQuestionSheet.requiredCol].toLowerCase() == 'y' || data[i][walkQuestionSheet.requiredCol].toLowerCase() == 'yes') {
                            walkQuestionSheet.chapterRQuestions[cqi].push(i);
                        } else {
                            walkQuestionSheet.chapterQuestions[cqi].push(i);
                        }
                    } else {
                        walkQuestionSheet.chapterQuestions[cqi].push(i);
                    }
                } else {
                    if (walkQuestionSheet.requiredCol > -1) {
                        if (data[i][walkQuestionSheet.requiredCol].toLowerCase() == 'true' || data[i][walkQuestionSheet.requiredCol].toLowerCase() == 't' || data[i][walkQuestionSheet.requiredCol].toLowerCase() == 'y' || data[i][walkQuestionSheet.requiredCol].toLowerCase() == 'yes') {
                            walkQuestionSheet.chapterRQuestions.push([i]);
                            walkQuestionSheet.chapterQuestions.push([]);
                        } else {
                            walkQuestionSheet.chapterRQuestions.push([]);
                            walkQuestionSheet.chapterQuestions.push([i]);
                        }
                    } else {
                        walkQuestionSheet.chapterRQuestions.push([]);
                        walkQuestionSheet.chapterQuestions.push([i]);
                    }

                    walkQuestionSheet.chapters.push(data[i][walkQuestionSheet.chapterCol]);
                    if (walkQuestionSheet.questionCountCol > -1) {
                        if (isInt(data[i][walkQuestionSheet.questionCountCol])) {
                            walkQuestionSheet.numberofquestions.push(data[i][walkQuestionSheet.questionCountCol]);
                        } else {
                            walkQuestionSheet.numberofquestions.push(walkQuestionSheet.numberofquestionsdefault);
                        }
                    } else {
                        walkQuestionSheet.numberofquestions.push(walkQuestionSheet.numberofquestionsdefault);
                    }
                }
            }
            for (var i = 0; i < walkQuestionSheet.chapters.length; i++) {
                var reqq = walkQuestionSheet.chapterRQuestions[i].length;
                var qcount = walkQuestionSheet.numberofquestions[i];
                walkQuestionSheet.numberofquestions[i] = qcount - reqq;
                if (walkQuestionSheet.chapterQuestions[i].length < walkQuestionSheet.numberofquestions[i]) {
                    walkQuestionSheet.numberofquestions[i] = walkQuestionSheet.chapterQuestions[i].length;
                }
            }
        } else {
            for (var i = 1; i < data.length; i++) {
                if (walkQuestionSheet.requiredCol > -1) {
                    if (data[i][walkQuestionSheet.requiredCol].toLowerCase() == 'true' || data[i][walkQuestionSheet.requiredCol].toLowerCase() == 't' || data[i][walkQuestionSheet.requiredCol].toLowerCase() == 'y' || data[i][walkQuestionSheet.requiredCol].toLowerCase() == 'yes') {
                        walkQuestionSheet.chapterRQuestions.push(i);
                    } else {
                        walkQuestionSheet.chapterQuestions.push(i);
                    }

                } else {
                    walkQuestionSheet.chapterQuestions.push(i);
                }
            }
            var reqq = walkQuestionSheet.chapterRQuestions.length;
            var qcount = walkQuestionSheet.numberofquestionsdefault;
            walkQuestionSheet.numberofquestions = qcount - reqq;
            if (walkQuestionSheet.chapterQuestions.length < walkQuestionSheet.numberofquestions) {
                walkQuestionSheet.numberofquestions = walkQuestionSheet.chapterQuestions.length;
            }
        }
        if (walkQuestionSheet.questionCol == -1) {
            return false;
        }
        return firstSheet;
    }
    return false;
};

newEmailClass = {};

newEmailClass.sendEmail = function(file, assignment, emailContent, studentName, studentEmail, subjectLine) {
    subjectLine = subjectLine || assignment;
    var url = file.getUrl();
    var content = studentName + "- \r\n\r\n";
    content += emailContent + "\r\n\r\n";
    content += url + "\r\n\r\n";
    if (interfaceClass.sendfile === true) {
        MailApp.sendEmail(studentEmail, subjectLine, content, {
            attachments: [file.getAs(MimeType.PDF)],
            name: file.getName()
        });
    } else {
        MailApp.sendEmail(studentEmail, subjectLine, content);
    }
};