document.addEventListener('DOMContentLoaded', function() {
    const fileInput = document.getElementById('fileInput');
    const emailsContainer = document.getElementById('emailsContainer');
  
    fileInput.addEventListener('change', processFiles);
  
    function processFiles(event) {
      const files = event.target.files;
      if (files.length === 0) {
        return;
      }
  
      const emailSet = new Set();
  
      for (let i = 0; i < files.length; i++) {
        const file = files[i];
        if (file.name.endsWith('.xlsx')) {
          const reader = new FileReader();
          reader.onload = function(event) {
            const arrayBuffer = event.target.result;
            const workbook = XLSX.read(new Uint8Array(arrayBuffer), { type: 'array' });
  
            workbook.SheetNames.forEach(function(sheetName) {
              const sheet = workbook.Sheets[sheetName];
              const sheetData = XLSX.utils.sheet_to_json(sheet);
              extractEmailsFromSheetData(sheetData, emailSet);
            });
  
            updateEmailsContainer(emailSet);
          };
  
          reader.readAsArrayBuffer(file);
        }
      }
    }
  
    function extractEmailsFromSheetData(sheetData, emailSet) {
      sheetData.forEach(row => {
        for (const key in row) {
          if (row.hasOwnProperty(key)) {
            const cellValue = row[key];
            if (typeof cellValue === 'string') {
              const emailMatches = cellValue.match(/([a-zA-Z0-9._-]+@[a-zA-Z0-9._-]+\.[a-zA-Z0-9._-]+)/gi);
              if (emailMatches) {
                emailMatches.forEach(email => emailSet.add(email));
              }
            }
          }
        }
      });
    }
  
    function updateEmailsContainer(emailSet) {
      emailsContainer.innerHTML = '';
      if (emailSet.size > 0) {
        const emailsList = document.createElement('ul');
        emailSet.forEach(email => {
          const listItem = document.createElement('li');
          listItem.textContent = email;
          emailsList.appendChild(listItem);
        });
        emailsContainer.appendChild(emailsList);
      } else {
        emailsContainer.textContent = 'No email addresses found.';
      }
    }
  });
  