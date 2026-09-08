function generateUniqueName(subject, index = null) {
  const now = new Date();

  // Date-Time format (DDMMYYHHmm)
  const dd = String(now.getDate()).padStart(2, "0");
  const mm = String(now.getMonth() + 1).padStart(2, "0");
  const yy = String(now.getFullYear()).slice(-2);
  const hh = String(now.getHours()).padStart(2, "0");
  const min = String(now.getMinutes()).padStart(2, "0");

  // Random 3-digit number for extra uniqueness
  const random = Math.floor(100 + Math.random() * 900);

  const currentTime = `${dd}${mm}${yy}${hh}${min}${random}`;

  return index !== undefined && index !== null
    ? `${subject}_${currentTime}_${index}`
    : `${subject}_${currentTime}`;
}


function createImageDoc(qnaList, folderId, subjectName = "Testing") {
  const folder = DriveApp.getFolderById(folderId);
  const imageUrls = [];

  // ✅ Create Google Doc first
  const doc = DocumentApp.create(`Doc_${generateUniqueName(subjectName)}`);
  const body = doc.getBody();

  qnaList.forEach((qna, index) => {
    let { questionText, selectedFile64String, selectedFileType, selectedFileName } = qna;

    // 📝 Add Question
    if (questionText) {
      body.appendParagraph("Q" + (index + 1) + ": " + questionText)
          .setHeading(DocumentApp.ParagraphHeading.HEADING2);
    }

    // 🖼️ If image present
    if (selectedFile64String) {
      if (selectedFile64String.indexOf("base64,") !== -1) {
        selectedFile64String = selectedFile64String.split("base64,")[1];
      }

      const binaryData = Utilities.base64Decode(selectedFile64String);
      const uniqueName = generateUniqueName(subjectName, index);
      const blob = Utilities.newBlob(binaryData, selectedFileType, uniqueName);

      // Save in Drive
      const newFile = folder.createFile(blob);
      newFile.setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.VIEW);
      imageUrls.push(`https://drive.google.com/file/d/${newFile.getId()}/view?usp=sharing`);

      // Insert into Doc
      const image = body.appendImage(blob);
      const pageWidth = body.getPageWidth() - body.getMarginLeft() - body.getMarginRight();
      const scaleFactor = pageWidth / image.getWidth();
      image.setWidth(pageWidth);
      image.setHeight(image.getHeight() * scaleFactor);
    }

    body.appendParagraph("-----------------------------");
  });

  // Move Doc to same folder
  const docFile = DriveApp.getFileById(doc.getId());
  folder.addFile(docFile);
  DriveApp.getRootFolder().removeFile(docFile);

  docFile.setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.VIEW);

  return {
    imageUrls: imageUrls,
    docUrl: doc.getUrl(),
  };
}

//api constant
const GET_PAPER_STRUCTURE = "GET_PAPER_STRUCTURE"
const TESTING_DOCS = "TESTING_DOCS"
const TESTING_IMAGEDOC = "TESTING_IMAGEDOC"
const TESTING_IMAGEDOC_MULTI = "TESTING_IMAGEDOC_MULTI"
const SHEET_ID = "1zVFRVL7L-6i10EKcHp1MZWoiSOXJ79iM6GTtAoVZ4yI";
const ARCHIEVE_SHEET_ID = "1fgBQZ6S9Cfu4JgeHZLmfdApkTnivz70ezMpyscTkw8U"

function doGet(request) {
  return HtmlService.createTemplateFromFile('Index')
    .evaluate().setSandboxMode(HtmlService.SandboxMode.IFRAME)
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
}

function include(filename) {
  return HtmlService.createHtmlOutputFromFile(filename)
    .getContent();
}


function doPost(e) {
  var data = JSON.parse(e.postData.contents);
  const { apiType, ...request } = data;

  const otherInfo = {
    apiType: apiType,
  }
  if (otherInfo.apiType === GET_PAPER_STRUCTURE) {
    
      try {
    const response = getPaperStructure();
    return ContentService.createTextOutput(JSON.stringify({ status: "success", data: response, request: data }))
        .setMimeType(ContentService.MimeType.JSON);


      }
      catch (ex) {
      return ContentService.createTextOutput(JSON.stringify({ status: "failuretest", data: ex.toString(), request: data }))
        .setMimeType(ContentService.MimeType.JSON);
    }
  
  }
  else if (otherInfo.apiType === TESTING_DOCS) {
    
      try {
    const { paperStructure } = request;
    const response = ssFileGenerate(paperStructure)
    return ContentService.createTextOutput(JSON.stringify({ status: "success", data: response, request: data }))
        .setMimeType(ContentService.MimeType.JSON);


      }
      catch (ex) {
      return ContentService.createTextOutput(JSON.stringify({ status: "failure", data: ex, request: data }))
        .setMimeType(ContentService.MimeType.JSON);
    }
  }
  else if (otherInfo.apiType === TESTING_IMAGEDOC) {
  try {
    let { selectedFileType, selectedFileName, selectedFile64String, ...updatedRequest } = request;

    var folderId = "1RM9bBgIPWGYvaL4gaqSHy6ag5MHc2niP";
    const folder = DriveApp.getFolderById(folderId);

    if (selectedFile64String.indexOf("base64,") !== -1) {
      selectedFile64String = selectedFile64String.split("base64,")[1];
    }

    const binaryData = Utilities.base64Decode(selectedFile64String);
    const blob = Utilities.newBlob(binaryData, selectedFileType, selectedFileName);
    const newFile = folder.createFile(blob);
    newFile.setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.VIEW);

    const imageUrl = `https://drive.google.com/file/d/${newFile.getId()}/view?usp=sharing`;

    // Step 2: Create Google Doc and insert image
    const doc = DocumentApp.create("Image Doc - " + selectedFileName);
    const body = doc.getBody();
    body.appendParagraph("Here is your uploaded image:");
    const image = body.appendImage(blob);

    const pageWidth = body.getPageWidth() - body.getMarginLeft() - body.getMarginRight();
    const originalWidth = image.getWidth();
    const originalHeight = image.getHeight();
    const scaleFactor = pageWidth / originalWidth;

    image.setWidth(pageWidth);
    image.setHeight(originalHeight * scaleFactor);

    // ✅ Move Doc to same folder
    const docFile = DriveApp.getFileById(doc.getId());
    folder.addFile(docFile);
    DriveApp.getRootFolder().removeFile(docFile); // remove from My Drive root

    // ✅ Set sharing to Anyone with link (View)
    docFile.setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.VIEW);

    const docUrl = doc.getUrl();

    const response = {
      imageUrl: imageUrl,
      docUrl: docUrl,
    };

    return ContentService.createTextOutput(JSON.stringify({ status: "success", data: response, request: data }))
      .setMimeType(ContentService.MimeType.JSON);
  }
    catch (ex) {
      return ContentService.createTextOutput(JSON.stringify({ status: "failure", data: ex.toString(), request: data }))
        .setMimeType(ContentService.MimeType.JSON);
    }
  }

  else if (otherInfo.apiType == TESTING_IMAGEDOC_MULTI)
  {
      try {
     var folderId = "1RM9bBgIPWGYvaL4gaqSHy6ag5MHc2niP";
      const response = createImageDoc(request.qnaList, folderId);
      return ContentService.createTextOutput(JSON.stringify({ status: "success", data: response, request: data }))
      .setMimeType(ContentService.MimeType.JSON);
      }
       catch (ex) {
      return ContentService.createTextOutput(JSON.stringify({ status: "failure", data: ex.toString(), request: data }))
        .setMimeType(ContentService.MimeType.JSON);
    }

  }
  else {
    return ContentService.createTextOutput(JSON.stringify({ status: "not found", status: false, request: data }))
      .setMimeType(ContentService.MimeType.JSON);
  }
};


function uploadDocFromHTML(request) {
  try {
    let imageFileUrl = "";
    if (request.selectedFile64String) {
      imageFileUrl = uploadFileToDrive(
        request.selectedFile64String,
        request?.selectedFileType,
        request?.selectedFileName || "UploadedImage.png"
      );
    }

    // ✅ Google Doc create karte hain
    const folderId = "1RM9bBgIPWGYvaL4gaqSHy6ag5MHc2niP"; // same folder
    const folder = DriveApp.getFolderById(folderId);

    const doc = DocumentApp.create(request.docTitle || "Uploaded Document");
    const body = doc.getBody();

    // Title
    body.appendParagraph(request.docTitle || "Untitled Document")
        .setHeading(DocumentApp.ParagraphHeading.HEADING1);

    // Question text
    if (request.questionText) {
      body.appendParagraph(request.questionText).setHeading(DocumentApp.ParagraphHeading.NORMAL);
    }

    // Insert Image from Base64 (not from Drive URL, directly decode again)
    if (request.selectedFile64String) {
      let base64String = request.selectedFile64String;
      if (base64String.indexOf("base64,") !== -1) {
        base64String = base64String.split("base64,")[1];
      }
      const binaryData = Utilities.base64Decode(base64String);
      const blob = Utilities.newBlob(binaryData, request.selectedFileType, request.selectedFileName);
      body.appendImage(blob).setWidth(300); // resize for doc
    }

    // Save Doc inside target folder
    const docFile = DriveApp.getFileById(doc.getId());
    folder.addFile(docFile);
    DriveApp.getRootFolder().removeFile(docFile); // remove from root

    // Sharing settings
    docFile.setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.VIEW);

    const docUrl = doc.getUrl();

    return {
      imageUrl: imageFileUrl, // drive image url (optional)
      docUrl: docUrl
    };

  } catch (ex) {
    console.log(ex.toString());
    return { error: ex.toString() };
  }
}


function uploadFileToDrive(base64String, fileType, fileName) {
   if (base64String.indexOf("base64,") !== -1) {
    base64String = base64String.split("base64,")[1];
  }


  // Decode Base64 to binary
  const binaryData = Utilities.base64Decode(base64String);

  // Create a new Blob with the decoded binary data
  const blob = Utilities.newBlob(binaryData, fileType, fileName);

  var folderId = "1RM9bBgIPWGYvaL4gaqSHy6ag5MHc2niP";
  // Get the target folder
  const folder = DriveApp.getFolderById(folderId);

  // Create the file in Google Drive
  const newFile = folder.createFile(blob);

  // Set sharing permissions to anyone with the link
  newFile.setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.VIEW);

  // Get the file ID
  const fileId = newFile.getId();

  // Construct the view URL
  const viewUrl = `https://drive.google.com/file/d/${fileId}/view?usp=sharing`;

  return viewUrl; // Return the view URL
}




function asdfasdtffGenerate() {
  const data = 
    {
    "apiType": "TESTING_DOCS",
    "subject": "English",
    "paperStructure" : [
    {
        "section": "Section A (Reading Skills)",
        "marks": 20,
        "questions": [
            {
                "type": "Unseen Passage 1",
                "questionHeading": "Read the following passage.",
                "count": 1,
                "marksPerQuestion": 10,
                "marks": 10,
                "no": 1,
                "text": "test 123"
            },
            {
                "type": "Unseen Passage 2",
                "questionHeading": "Read the following passage.",
                "count": 1,
                "marksPerQuestion": 10,
                "marks": 10,
                "no": 2,
                "text": "test 123"
            }
        ]
    },
    {
        "section": "Section B (Writing & Grammar)",
        "marks": 20,
        "questions": [
            {
                "type": "Grammar",
                "questionHeading": "Complete any ten of twelve of the following tasks, as directed.",
                "count": 12,
                "marksPerQuestion": 1,
                "marks": 10,
                "choice": "Attempt any 10 out of 12",
                "no": 1,
                "text": "test 123"
            },
            {
                "type": "Grammar",
                "questionHeading": "Complete any ten of twelve of the following tasks, as directed.",
                "count": 12,
                "marksPerQuestion": 1,
                "marks": 10,
                "choice": "Attempt any 10 out of 12",
                "no": 2,
                "text": "test 123"
            },
            {
                "type": "Grammar",
                "questionHeading": "Complete any ten of twelve of the following tasks, as directed.",
                "count": 12,
                "marksPerQuestion": 1,
                "marks": 10,
                "choice": "Attempt any 10 out of 12",
                "no": 3,
                "text": "test 123"
            },
            {
                "type": "Grammar",
                "questionHeading": "Complete any ten of twelve of the following tasks, as directed.",
                "count": 12,
                "marksPerQuestion": 1,
                "marks": 10,
                "choice": "Attempt any 10 out of 12",
                "no": 4,
                "text": "test 123"
            },
            {
                "type": "Grammar",
                "questionHeading": "Complete any ten of twelve of the following tasks, as directed.",
                "count": 12,
                "marksPerQuestion": 1,
                "marks": 10,
                "choice": "Attempt any 10 out of 12",
                "no": 5,
                "text": "test 123"
            },
            {
                "type": "Grammar",
                "questionHeading": "Complete any ten of twelve of the following tasks, as directed.",
                "count": 12,
                "marksPerQuestion": 1,
                "marks": 10,
                "choice": "Attempt any 10 out of 12",
                "no": 6,
                "text": "test 123"
            },
            {
                "type": "Grammar",
                "questionHeading": "Complete any ten of twelve of the following tasks, as directed.",
                "count": 12,
                "marksPerQuestion": 1,
                "marks": 10,
                "choice": "Attempt any 10 out of 12",
                "no": 7,
                "text": "asdfasdfasdf"
            },
            {
                "type": "Grammar",
                "questionHeading": "Complete any ten of twelve of the following tasks, as directed.",
                "count": 12,
                "marksPerQuestion": 1,
                "marks": 10,
                "choice": "Attempt any 10 out of 12",
                "no": 8,
                "text": "asdfasdfasdfasdf"
            },
            {
                "type": "Grammar",
                "questionHeading": "Complete any ten of twelve of the following tasks, as directed.",
                "count": 12,
                "marksPerQuestion": 1,
                "marks": 10,
                "choice": "Attempt any 10 out of 12",
                "no": 9,
                "text": "we34234234234"
            },
            {
                "type": "Grammar",
                "questionHeading": "Complete any ten of twelve of the following tasks, as directed.",
                "count": 12,
                "marksPerQuestion": 1,
                "marks": 10,
                "choice": "Attempt any 10 out of 12",
                "no": 10,
                "text": "asdfasdfasdf"
            },
            {
                "type": "Grammar",
                "questionHeading": "Complete any ten of twelve of the following tasks, as directed.",
                "count": 12,
                "marksPerQuestion": 1,
                "marks": 10,
                "choice": "Attempt any 10 out of 12",
                "no": 11,
                "text": "asdfasdfasdf"
            },
            {
                "type": "Grammar",
                "questionHeading": "Complete any ten of twelve of the following tasks, as directed.",
                "count": 12,
                "marksPerQuestion": 1,
                "marks": 10,
                "choice": "Attempt any 10 out of 12",
                "no": 12,
                "text": "asdfasdfasdf"
            },
            {
                "type": "Writing Q4",
                "questionHeading": "All details presented in the questions in writing section are imaginary and created for assessment purpose.",
                "count": 2,
                "marksPerQuestion": 5,
                "marks": 5,
                "choice": "Do any 1 out of 2",
                "no": 13,
                "text": "asdfasdfasdf"
            },
            {
                "type": "Writing Q4",
                "questionHeading": "All details presented in the questions in writing section are imaginary and created for assessment purpose.",
                "count": 2,
                "marksPerQuestion": 5,
                "marks": 5,
                "choice": "Do any 1 out of 2",
                "no": 14,
                "text": "asdfasdfasdf"
            },
            {
                "type": "Writing Q5",
                "questionHeading": "All details presented in the questions in writing section are imaginary and created for assessment purpose.",
                "count": 2,
                "marksPerQuestion": 5,
                "marks": 5,
                "choice": "Do any 1 out of 2",
                "no": 15,
                "text": "asdfasdfasdf"
            },
            {
                "type": "Writing Q5",
                "questionHeading": "All details presented in the questions in writing section are imaginary and created for assessment purpose.",
                "count": 2,
                "marksPerQuestion": 5,
                "marks": 5,
                "choice": "Do any 1 out of 2",
                "no": 16,
                "text": "asdfasdfasrqwrqweras\nfa\nsdf\nas\ndf\nasdf"
            }
        ]
    },
    {
        "section": "Section C (Literature Book)",
        "marks": 35,
        "questions": [
            {
                "type": "Extract Q6",
                "questionHeading": "Read the given extracts A and B and answer ANY ONE of the two.",
                "count": 2,
                "marksPerQuestion": 5,
                "marks": 5,
                "choice": "Attempt any 1 out of 2",
                "no": 1,
                "text": "asdfasdf"
            },
            {
                "type": "Extract Q6",
                "questionHeading": "Read the given extracts A and B and answer ANY ONE of the two.",
                "count": 2,
                "marksPerQuestion": 5,
                "marks": 5,
                "choice": "Attempt any 1 out of 2",
                "no": 2,
                "text": "asdfasdf"
            },
            {
                "type": "Extract Q7",
                "questionHeading": "Read the given extracts A and B and answer ANY ONE of the two",
                "count": 2,
                "marksPerQuestion": 5,
                "marks": 5,
                "choice": "Attempt any 1 out of 2",
                "no": 3,
                "text": "asdfasdf"
            },
            {
                "type": "Extract Q7",
                "questionHeading": "Read the given extracts A and B and answer ANY ONE of the two",
                "count": 2,
                "marksPerQuestion": 5,
                "marks": 5,
                "choice": "Attempt any 1 out of 2",
                "no": 4,
                "text": "asdfasdfa"
            },
            {
                "type": "Short Answer Q8",
                "questionHeading": "Answer ANY FOUR of the following five questions, in about 50 words each.",
                "count": 5,
                "marksPerQuestion": 3,
                "marks": 12,
                "choice": "Attempt any 4 out of 5",
                "no": 5,
                "text": "asdfasdfas"
            },
            {
                "type": "Short Answer Q8",
                "questionHeading": "Answer ANY FOUR of the following five questions, in about 50 words each.",
                "count": 5,
                "marksPerQuestion": 3,
                "marks": 12,
                "choice": "Attempt any 4 out of 5",
                "no": 6,
                "text": "dfasdfasdfasdf"
            },
            {
                "type": "Short Answer Q8",
                "questionHeading": "Answer ANY FOUR of the following five questions, in about 50 words each.",
                "count": 5,
                "marksPerQuestion": 3,
                "marks": 12,
                "choice": "Attempt any 4 out of 5",
                "no": 7,
                "text": "asdfasdfasdf"
            },
            {
                "type": "Short Answer Q8",
                "questionHeading": "Answer ANY FOUR of the following five questions, in about 50 words each.",
                "count": 5,
                "marksPerQuestion": 3,
                "marks": 12,
                "choice": "Attempt any 4 out of 5",
                "no": 8,
                "text": "asdfasdfasdf"
            },
            {
                "type": "Short Answer Q8",
                "questionHeading": "Answer ANY FOUR of the following five questions, in about 50 words each.",
                "count": 5,
                "marksPerQuestion": 3,
                "marks": 12,
                "choice": "Attempt any 4 out of 5",
                "no": 9,
                "text": "asdfasdfasdf"
            },
            {
                "type": "Short Answer Q9",
                "questionHeading": "Answer ANY TWO of the following three questions, in about 40-50 words.",
                "count": 3,
                "marksPerQuestion": 3,
                "marks": 6,
                "choice": "Attempt any 2 out of 3",
                "no": 10,
                "text": "asdfasdfasdf"
            },
            {
                "type": "Short Answer Q9",
                "questionHeading": "Answer ANY TWO of the following three questions, in about 40-50 words.",
                "count": 3,
                "marksPerQuestion": 3,
                "marks": 6,
                "choice": "Attempt any 2 out of 3",
                "no": 11,
                "text": "asdfasdfasdfasdf"
            },
            {
                "type": "Short Answer Q9",
                "questionHeading": "Answer ANY TWO of the following three questions, in about 40-50 words.",
                "count": 3,
                "marksPerQuestion": 3,
                "marks": 6,
                "choice": "Attempt any 2 out of 3",
                "no": 12,
                "text": "asdfasdfasdf"
            },
            {
                "type": "Long Answer Q10",
                "questionHeading": "Answer ANY ONE of the following two questions, in about 100-120 words",
                "count": 2,
                "marksPerQuestion": 6,
                "marks": 6,
                "choice": "Attempt any 1 out of 2",
                "no": 13,
                "text": "asdfasdfasdf"
            },
            {
                "type": "Long Answer Q10",
                "questionHeading": "Answer ANY ONE of the following two questions, in about 100-120 words",
                "count": 2,
                "marksPerQuestion": 6,
                "marks": 6,
                "choice": "Attempt any 1 out of 2",
                "no": 14,
                "text": "asdfasdfasdf\nasd\nfa\nsdf"
            },
            {
                "type": "Long Answer Q11",
                "questionHeading": "Answer ANY ONE of the following two questions, in about 100-120 words.",
                "count": 2,
                "marksPerQuestion": 6,
                "marks": 6,
                "choice": "Attempt any 1 out of 2",
                "no": 15,
                "text": "asdfasdfasd"
            },
            {
                "type": "Long Answer Q11",
                "questionHeading": "Answer ANY ONE of the following two questions, in about 100-120 words.",
                "count": 2,
                "marksPerQuestion": 6,
                "marks": 6,
                "choice": "Attempt any 1 out of 2",
                "no": 16,
                "text": "asdfasdfasdfasdf"
            }
        ]
    }
]


  }

  const { paperStructure } = data;
  console.log(ssFileGenerate(paperStructure))
}

function ssFileGenerate(paperStructure) {
  try {
    const doc = DocumentApp.create("Generated Question Paper");
    const body = doc.getBody();

    // Title
    body.appendParagraph("GAURANG GURUKUL QUESTION PAPER")
      .setHeading(DocumentApp.ParagraphHeading.HEADING1)
      .setAlignment(DocumentApp.HorizontalAlignment.CENTER);
    body.appendParagraph(" ");

    paperStructure.forEach(section => {
      // Section Title
      body.appendParagraph(`${section.section} [${section.marks} Marks]`)
        .setHeading(DocumentApp.ParagraphHeading.HEADING2)
        .setAlignment(DocumentApp.HorizontalAlignment.CENTER);
      body.appendParagraph(" ");

      // Group by TYPE only
      const choiceGroups = {};
      section.questions.forEach(q => {
        if (q.choice) {
          const key = q.type; // 👈 only type is used now
          if (!choiceGroups[key]) choiceGroups[key] = { heading: q.questionHeading, choice: q.choice, marks: q.marks, marksPerQuestion: q.marksPerQuestion, questions: [] };
          choiceGroups[key].questions.push(q);
        } else {
          // Questions without choice
          const noOfQuestions = q.marks / q.marksPerQuestion;
          const marksPara = body.appendParagraph(`(${noOfQuestions} * ${q.marksPerQuestion} = ${q.marks})`);
          marksPara.setAlignment(DocumentApp.HorizontalAlignment.RIGHT).setBold(true);

          const para = body.appendParagraph(`Q${q.no}: ${q.text || "[No text provided]"}`);
          para.setSpacingAfter(8);
        }
      });

      // Process grouped questions
      Object.keys(choiceGroups).forEach(type => {
        const group = choiceGroups[type];
        const noOfQuestions = group.marks / group.marksPerQuestion;
        const marksDisplay = `(${noOfQuestions} * ${group.marksPerQuestion} = ${group.marks})`;

       // Heading
  body.appendParagraph(group.heading).setBold(true);

  // Choice
  body.appendParagraph(group.choice).setBold(true);

  // Marks (right aligned ✅)
  body.appendParagraph(marksDisplay)
    .setAlignment(DocumentApp.HorizontalAlignment.RIGHT)
    .setBold(true);

        group.questions.forEach(q => {
          const para = body.appendParagraph(`Q${q.no}: ${q.text || "[No text provided]"}`);
          para.setSpacingAfter(8);
        });

        body.appendParagraph(" ");
      });

      body.appendParagraph(" ");
    });


       DriveApp.getFileById(doc.getId()).setSharing(
      DriveApp.Access.ANYONE_WITH_LINK,
      DriveApp.Permission.VIEW
    );
    return { docUrl: doc.getUrl() };
  } catch (ex) {
    return ex.toString();
  }
}



function getPaperStructure(){
const sheetMaster = SpreadsheetApp.openByUrl("https://docs.google.com/spreadsheets/d/1PGYZ6NW7C3IqLFiAJyvnFuzaY_6pxRcsoWUnZJGh6u4/edit?gid=1552451432#gid=1552451432").getSheetByName("Master");
  const data = sheetMaster.getDataRange().getValues();

  // Header row
  const headers = data[0];
  const rows = data.slice(1);

  // Group by subject → section
  const subjectMap = {};

  rows.forEach(row => {
    const rowObj = {};
    headers.forEach((h, i) => rowObj[h] = row[i]);

    const subjectName = rowObj["Subject"];
    const sectionCode = rowObj["Section"];
    const sectionTitle = rowObj["Section Title"];
    const questionType = rowObj["Question Type"];
    const questionHeading = rowObj["Question Heading"];
    const count = Number(rowObj["No. of Questions"]) || 0;
    const marksPerQ = Number(rowObj["Marks/Questions"]) || 0;
    const marks = Number(rowObj["Marks"]) || 0;
    const choice = rowObj["Choice Type"] || "";

    if (!subjectMap[subjectName]) {
      subjectMap[subjectName] = {};
    }

    if (!subjectMap[subjectName][sectionTitle]) {
      subjectMap[subjectName][sectionTitle] = {
        name: sectionTitle,
        marks: 0,
        questions: []
      };
    }

    // ✅ Add question
    subjectMap[subjectName][sectionTitle].questions.push({
      type: questionType,
      questionHeading: questionHeading,
      count: count,
      marksPerQuestion: marksPerQ,
      marks: marks,
      choice: choice
    });

    // ✅ Add section marks
    subjectMap[subjectName][sectionTitle].marks += marks;
  });

  // Convert subjectMap → final JSON array
  const finalOutput = Object.keys(subjectMap).map(subjectName => {
    return {
      subject: subjectName,
      sections: Object.values(subjectMap[subjectName])
    };
  });

  console.log(JSON.stringify(finalOutput, null, 2));
  return finalOutput
}

function asdtesting() {
  uploadDocFromHTML()
}


<script>
document.addEventListener("DOMContentLoaded", () => {
  loadPaperStructure();
});

const subjectSelect = document.getElementById("subjectSelect");
let paperStructure = []; //  API response store

function renderForm(subjectIndex) {
  const formEl = document.getElementById("paperForm");
  if (!formEl) {
    console.error("❌ formContainer element not found in DOM");
    return;
  }

  formEl.innerHTML = ""; //  clear old content
  const subject = paperStructure[subjectIndex];
  if (!subject) {
    console.error("❌ Invalid subject index:", subjectIndex);
    return;
  }

  let questionCounter = 1; //  Global counter for numbering

  subject.sections.forEach((section, si) => {
    const secDiv = document.createElement("div");
    secDiv.innerHTML = `<h3>${section.name} [${section.marks} Marks]</h3>`;

    section.questions.forEach((q, qi) => {
      const typeBlock = document.createElement("div");
      typeBlock.className = "question-type-block";
      let effectiveCount = Math.round(q.marks / q.marksPerQuestion);

      const typeHeading = document.createElement("h4");
      typeHeading.innerHTML = `${q.type} (${q.marksPerQuestion} × ${effectiveCount} = ${q.marks})`;
      typeBlock.appendChild(typeHeading);

      for (let i = 0; i < q.count; i++) {
        const qDiv = document.createElement("div");
        qDiv.className = "question-item";
        qDiv.innerHTML = `
          <label>Q${questionCounter} [${q.marksPerQuestion} Marks]</label><br>
          ${q.questionHeading ? `<small>${q.questionHeading}</small><br>` : ""}
          ${q.choice ? `<div class="choice">${q.choice}</div>` : ""}
          <textarea name="s${si}_q${qi}_c${i}" rows="2" style="width:100%;" placeholder="Enter question text"></textarea>
          <br>
          <input type="file" 
                 name="img_s${si}_q${qi}_c${i}" 
                 accept="image/*" />
          <br><br>
        `;
        typeBlock.appendChild(qDiv);
        questionCounter++;
      }

      secDiv.appendChild(typeBlock);
    });

    formEl.appendChild(secDiv);
  });
}

function submitPaper() {
  const answers = [];
  const subjectIndex = subjectSelect.value;
  const subject = paperStructure[subjectIndex];

  const fileReaders = [];
  const promises = [];

  subject.sections.forEach((section, si) => {
    const secData = {
      section: section.name,
      marks: section.marks,
      questions: [],
    };
    let questionCounter = 1; // section-wise numbering

    section.questions.forEach((q, qi) => {
      for (let i = 0; i < q.count; i++) {
        const textValue = document.querySelector(
          `[name="s${si}_q${qi}_c${i}"]`
        ).value;

        const fileInput = document.querySelector(
          `[name="img_s${si}_q${qi}_c${i}"]`
        );
        const file = fileInput?.files?.[0];

        if (file) {
          const reader = new FileReader();
          const promise = new Promise((resolve) => {
            reader.onload = function (e) {
              secData.questions.push({
                ...q,
                no: questionCounter,
                text: textValue,
                imageBase64: e.target.result, // ✅ image bhi add ho gayi
              });
              resolve();
            };
          });
          reader.readAsDataURL(file);
          promises.push(promise);
        } else {
          secData.questions.push({
            ...q,
            no: questionCounter,
            text: textValue,
            imageBase64: null,
          });
        }

        questionCounter++;
      }
    });

    answers.push(secData);
    console.log("Section data final :", secData);
  });

  Promise.all(promises).then(() => {
    const saveRequest = {
      apiType: "TESTING_DOCS",
      subject: subject.subject,
      paperStructure: answers,
    };
    IsLoading(true);
    fetch(
      "https://script.google.com/macros/s/AKfycbztRT61BtnaLpEOcE9nghNdMz5rm-l8Oi8xbNhEO50V9TvrZBNtR_JE9h6hJxSqurx-Dw/exec",
      {
        method: "POST",
        body: JSON.stringify(saveRequest),
      }
    )
      .then((res) => res.json())
      .then((data) => {
        IsLoading(false);
        if (data.status === "success") {
          alert("Document Generated Successfully");
          document.getElementById(
            "preview"
          ).innerHTML = `<p style="color:green;">✅ Success</p>
             <p><a href="${data.data.docUrl}" target="_blank">📄 View Generated Paper</a></p>`;
        }
      });
  });
}


async function loadPaperStructure() {
  try {
    const res = await CALL_API("GET_PAPER_STRUCTURE", "");
    if (res && res.data) {
      paperStructure = res.data; // API se array milega
      populateSubjectDropdown();
      renderForm(0); // Default render first subject
    } else {
      SHOW_ERROR_POPUP("No data received from API");
    }
  } catch (err) {
    SHOW_ERROR_POPUP("API Error: " + err);
  }
}

function populateSubjectDropdown() {
  subjectSelect.innerHTML = "";
  paperStructure.forEach((subject, index) => {
    const option = document.createElement("option");
    option.value = index;
    option.textContent = subject.subject; // 👈 subject ka naam
    subjectSelect.appendChild(option);
  });
  subjectSelect.addEventListener("change", (e) => {
    renderForm(e.target.value);
  });
}

  </script>


  <!DOCTYPE html>
<html>
  <head>
    <title>Upload Multiple QnA with Images</title>
    <style>
      body {
        font-family: Arial, sans-serif;
        margin: 20px;
      }
      .container {
        max-width: 700px;
        margin: auto;
      }
      .qna-block {
        border: 1px solid #ccc;
        padding: 15px;
        margin-bottom: 20px;
        border-radius: 10px;
        background: #f9f9f9;
      }
      textarea,
      input,
      button {
        width: 100%;
        margin: 8px 0;
        padding: 8px;
      }
      .preview {
        margin-top: 10px;
        text-align: center;
      }
      img {
        max-width: 100%;
        border: 1px solid #ddd;
        border-radius: 6px;
        margin-top: 8px;
      }
      .result-box {
        margin-top: 20px;
        padding: 15px;
        border: 1px solid #ddd;
        border-radius: 10px;
        background: #eef;
      }
    </style>
  </head>
  <body>
    <div class="container">
      <h2>Upload Multiple QnA → Save to Drive & Docs</h2>

      <!-- Block 1 -->
      <div class="qna-block" data-index="1">
        <h3>Question 1</h3>
        <textarea placeholder="Type Question 1"></textarea>
        <input type="file" accept="image/*" onchange="previewFile(this)" />
        <div class="preview"></div>
      </div>

      <!-- Block 2 -->
      <div class="qna-block" data-index="2">
        <h3>Question 2</h3>
        <textarea placeholder="Type Question 2"></textarea>
        <input type="file" accept="image/*" onchange="previewFile(this)" />
        <div class="preview"></div>
      </div>

      <!-- Block 3 -->
      <div class="qna-block" data-index="3">
        <h3>Question 3</h3>
        <textarea placeholder="Type Question 3"></textarea>
        <input type="file" accept="image/*" onchange="previewFile(this)" />
        <div class="preview"></div>
      </div>

      <button onclick="uploadData()">Upload All & Create Doc</button>

      <div id="results"></div>
    </div>

    <script>
      let qnaData = {};

      function previewFile(inputEl) {
        const block = inputEl.closest(".qna-block");
        const index = block.getAttribute("data-index");
        const file = inputEl.files[0];

        if (file) {
          const reader = new FileReader();
          reader.onload = function (event) {
            qnaData[index] = qnaData[index] || {};
            qnaData[index].imageBase64 = event.target.result;
            qnaData[index].fileType = file.type;
            qnaData[index].fileName = file.name;

            block.querySelector(
              ".preview"
            ).innerHTML = `<img src="${event.target.result}" alt="Preview">`;
          };
          reader.readAsDataURL(file);
        }
      }

      function uploadData() {
        // collect all blocks data
        document.querySelectorAll(".qna-block").forEach((block) => {
          const index = block.getAttribute("data-index");
          const text = block.querySelector("textarea").value.trim();

          if (text) {
            qnaData[index] = qnaData[index] || {};
            qnaData[index].questionText = text;
          }
        });

        // convert object → array
        const qnaArray = Object.keys(qnaData).map((k) => ({
          questionText: qnaData[k].questionText,
          selectedFileType: qnaData[k].fileType,
          selectedFileName: qnaData[k].fileName,
          selectedFile64String: qnaData[k].imageBase64,
        }));

        if (qnaArray.length === 0) {
          alert("Please add at least one question with image!");
          return;
        }

        const saveRequest = {
          apiType: "TESTING_IMAGEDOC_MULTI",
          docTitle: "Multiple QnA Test Doc",
          qnaList: qnaArray,
        };

        fetch(
          "https://script.google.com/macros/s/AKfycbztRT61BtnaLpEOcE9nghNdMz5rm-l8Oi8xbNhEO50V9TvrZBNtR_JE9h6hJxSqurx-Dw/exec",
          {
            method: "POST",
            body: JSON.stringify(saveRequest),
          }
        )
          .then((res) => res.json())
          .then((data) => {
            if (data.status === "success") {
              document.getElementById("results").innerHTML = `
                <div class="result-box">
                  <p><b>✅ Document Created:</b> <a href="${data.docUrl}" target="_blank">Open Google Doc</a></p>
                </div>`;
            } else {
              alert("Upload failed!");
            }
          })
          .catch((err) => console.error(err));
      }
    </script>
  </body>
</html>


document.addEventListener("DOMContentLoaded", () => {
  loadPaperStructure();
});

const subjectSelect = document.getElementById("subjectSelect");
let paperStructure = []; // ✅ API response store

function renderForm(subjectIndex) {
  const formEl = document.getElementById("paperForm");
  if (!formEl) {
    console.error("❌ formContainer element not found in DOM");
    return;
  }

  formEl.innerHTML = ""; // ✅ clear old content
  const subject = paperStructure[subjectIndex];
  if (!subject) {
    console.error("❌ Invalid subject index:", subjectIndex);
    return;
  }

  let questionCounter = 1; // ✅ Global counter for numbering

  subject.sections.forEach((section, si) => {
    const secDiv = document.createElement("div");
    secDiv.innerHTML = `<h3>${section.name} [${section.marks} Marks]</h3>`;

    section.questions.forEach((q, qi) => {
      const typeBlock = document.createElement("div");
      typeBlock.className = "question-type-block";
      let effectiveCount = Math.round(q.marks / q.marksPerQuestion);

      const typeHeading = document.createElement("h4");
      typeHeading.innerHTML = `${q.type} (${q.marksPerQuestion} × ${effectiveCount} = ${q.marks})`;
      typeBlock.appendChild(typeHeading);

      for (let i = 0; i < q.count; i++) {
        const qDiv = document.createElement("div");
        qDiv.className = "question-item";
        qDiv.innerHTML = `
          <label>Q${questionCounter} [${q.marksPerQuestion} Marks]</label><br>
          ${q.questionHeading ? `<small>${q.questionHeading}</small><br>` : ""}
          ${q.choice ? `<div class="choice">${q.choice}</div>` : ""}
          <textarea name="s${si}_q${qi}_c${i}" rows="2" style="width:100%;" placeholder="Enter question text"></textarea>
          <br><br>
        `;
        typeBlock.appendChild(qDiv);
        questionCounter++;
      }

      secDiv.appendChild(typeBlock);
    });

    formEl.appendChild(secDiv);
  });
}

function submitPaper() {
  const answers = [];
  const subjectIndex = subjectSelect.value;
  const subject = paperStructure[subjectIndex];

  subject.sections.forEach((section, si) => {
    const secData = {
      section: section.name,
      marks: section.marks,
      questions: [],
    };
    let questionCounter = 1; // section-wise numbering

    section.questions.forEach((q, qi) => {
      for (let i = 0; i < q.count; i++) {
        const ans = document.querySelector(
          `[name="s${si}_q${qi}_c${i}"]`
        ).value;
        secData.questions.push({ ...q, no: questionCounter, text: ans });
        questionCounter++;
      }
    });

    answers.push(secData);
  });

  const saveRequest = {
    apiType: "TESTING_DOCS",
    subject: subject.subject,
    paperStructure: answers,
  };
  IsLoading(true);
  fetch(
    "https://script.google.com/macros/s/AKfycbztRT61BtnaLpEOcE9nghNdMz5rm-l8Oi8xbNhEO50V9TvrZBNtR_JE9h6hJxSqurx-Dw/exec",
    {
      method: "POST",
      body: JSON.stringify(saveRequest),
    }
  )
    .then((res) => res.json())
    .then((data) => {
      IsLoading(false);
      if (data.status === "success") {
        alert("Document Generated Successfully");
        document.getElementById(
          "preview"
        ).innerHTML = `<p style="color:green;">✅ Success</p>
   <p><a href="${data.data.docUrl}" target="_blank">📄 View Generated Paper</a></p>`;
      }
    });
}

async function loadPaperStructure() {
  try {
    const res = await CALL_API("GET_PAPER_STRUCTURE", "");
    if (res && res.data) {
      paperStructure = res.data; // API se array milega
      populateSubjectDropdown();
      renderForm(0); // Default render first subject
    } else {
      SHOW_ERROR_POPUP("No data received from API");
    }
  } catch (err) {
    SHOW_ERROR_POPUP("API Error: " + err);
  }
}

function populateSubjectDropdown() {
  subjectSelect.innerHTML = "";
  paperStructure.forEach((subject, index) => {
    const option = document.createElement("option");
    option.value = index;
    option.textContent = subject.subject; // 👈 subject ka naam
    subjectSelect.appendChild(option);
  });
  subjectSelect.addEventListener("change", (e) => {
    renderForm(e.target.value);
  });
}
