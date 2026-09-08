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
          <input 
                class="ssPaperFileInput"
                type="file" 
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
