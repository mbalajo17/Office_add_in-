
/* eslint-disable office-addins/load-object-before-read */
/* eslint-disable office-addins/call-sync-before-read */
/* eslint-disable @typescript-eslint/no-unused-vars */
/* eslint-disable no-undef */
/*
 * Copyright (c) Microsoft Corporation. All rights reserved. Licensed under the MIT license.
 * See LICENSE in the project root for license information.
 */

/* global document, Office, Word */
let template; // Declare template as a global variable
Office.onReady((info) => {
  getFileName();
  if (info.host === Office.HostType.Word) {
    document.getElementById("sideload-msg").style.display = "none";
    document.getElementById("app-body").style.display = "flex";
    // document.getElementById("viewDocx").addEventListener("click", viewDocxAPI);
    document.getElementById("load-from-api").addEventListener("click", loadFromApi);
    document.getElementById("get-font-style").addEventListener("click", getFontStyle);
    document.getElementById("get-xml").addEventListener("click", getXML);
    // document.getElementById("get-textfor").addEventListener("click", getTeseData);
    
    document.getElementById("remove-doc").addEventListener("click", removeDoc);
    document.getElementById("restrict-section").addEventListener("click", restrictSection);
    document.getElementById("apiButton").addEventListener("click", genrateCSR)
    document.getElementById("submit").addEventListener("click", uplaodProtocal)
    document.getElementById("sapsubmit").addEventListener("click", uplaodSAP)
    document.getElementById("intext").addEventListener("click", uplaodINTEXT)

    document.getElementById("save-button").addEventListener("click", saveDocumentsss);
    document.getElementById("removeRestrictions").onclick = removeDocumentProtection;

  }
});


const bearerToken =
  "eyJ1c2VyIjp7ImlkIjoxMiwidXNlck5hbWUiOiJiYWxhamltb2hhbkBzeW1iaWFuY2UuY29tIiwicm9sZXMiOnsiaWQiOjMsInJvbGVOYW1lIjoiVXNlciJ9LCJhdXRob3JpdGllcyI6bnVsbCwiZmlyc3ROYW1lIjoiYmFsYWppIiwibGFzdE5hbWUiOiJtb2hhbiIsIm1pZGRsZU5hbWUiOm51bGwsInNob3dIZWxwIjp0cnVlLCJvcmdhbml6YXRpb24iOnsiaWQiOjEsIm9yZ05hbWUiOiJTeW1iaWFuY2UifX0sImFsZyI6IkhTNTEyIn0.eyJzdWIiOiJiYWxhamltb2hhbkBzeW1iaWFuY2UuY29tIiwiYXV0aG9yaXRpZXMiOlsiVXNlciJdLCJpYXQiOjE3MjUyNzc5NzEsImV4cCI6MTcyNTM2NDM3MX0.ruxwt7i6auSwtRd5hU2651EVL3EgW7KFpfQ6j0_Wz6georf8_upwr7RDSmnJ4tt1nf-JjInCU9Ox0AQOnVRVMw";

const uploadStatusprotocol = document.getElementById("uploadStatusprotocol");
const uploadStatusSAP = document.getElementById("uploadStatusSAP");
const uploadStatusINTEXT = document.getElementById("uploadStatusINTEXT");


async function restrictSection() {
  await Word.run(async function (context) {
    const paragraphs = context.document.body.paragraphs;
    paragraphs.load("items");
    await context.sync();

    await headingTextRestric(context, paragraphs, "Title Page", "Table of Contents");
    await headingTextRestric(context, paragraphs, "Introduction", "Study Objectives");
  });
}

async function headingTextRestric(context, paragraphs, startText, endText) {
  let startIndex = -1;
  let endIndex = -1;

  for (let i = 0; i < paragraphs.items.length; i++) {
    const paragraph = paragraphs.items[i];
    paragraph.load("text");
    await context.sync();

    if (paragraph.text.trim() === startText && startIndex === -1) {
      startIndex = i;
    } else if (paragraph.text.trim() === endText && startIndex !== -1) {
      endIndex = i - 1; // Changed from i-2 to i-1
      break;
    }
  }

  if (startIndex !== -1 && endIndex !== -1) {
    const startParagraph = paragraphs.items[startIndex];
    const endParagraph = paragraphs.items[endIndex];
    const range = startParagraph.getRange().expandTo(endParagraph.getRange());

    const contentControl = range.insertContentControl();
    contentControl.cannotEdit = true;
    contentControl.title = "Restricted Section";
    contentControl.tag = "RestrictedSection";

    await context.sync();
    console.log("Restricted section applied successfully.");
  } else {
    console.log(`No section found between "${startText}" and "${endText}".`);
  }
}

async function removeDocumentProtection() {
  Word.run(context => {
    const doc = context.document;
    const contentControls = doc.contentControls;

    contentControls.load("items");
    return context.sync().then(() => {
      for (let i = 0; i < contentControls.items.length; i++) {
        contentControls.items[i].delete(true); // true to delete the content inside the content control
      }
    }).then(context.sync);
  }).catch(error => {
    console.error("Error: " + error);
  });
}

function genrateCSR() {
  // Define the API endpoint URL
  const studyId = 1893;
  const apiUrl = "http://localhost:9156/csrfile/process/" + studyId;

  // Define the payload data
  const payload = {
    safetyRequired: false,
    tenseRequired: false,
    summaryRequired: false,
    detailInterpretationRequired: false,
    statChoice: "",
    enableTableReference: false,
    synopsisRequired: false,
    uploadedSynopsis: false,
  };

  // Make a POST request to the API
  // eslint-disable-next-line no-undef
  return fetch(apiUrl, {
    method: "POST",
    headers: {
      "Content-Type": "application/json",
      Authorization: "Bearer " + bearerToken,
    },
    body: JSON.stringify(payload),
  })
    .then((response) => {
      if (response.ok) {
        console.log("API call Successfull");
      } else {
        throw new Error("Failed to call the API");
      }
    })
    .then((data) => {
      // Handle the API response data here
      console.log("API response:", data);
    })
    .catch((error) => {
      console.error("API request error:", error);
    });
}

async function uplaodSAP(e) {
  e.preventDefault(); // Prevent the default form submission behavior

  // Get the selected file from the input field
  const fileInput = document.getElementById("sapInput")
  const file = fileInput.files[0];

  if (file) {
    try {
      const formData = new FormData();
      formData.append("file", file);

      const studyId = 1893;

      const uploadUrl = "http://localhost:9156/csrfile/SAP/upload/" + studyId;

      var headers = new Headers({
        Authorization: "Bearer " + bearerToken,
      });

      const response = await fetch(uploadUrl, {
        method: "POST",
        headers: headers,
        body: formData,
      });

      if (response.ok) {
        uploadStatusSAP.textContent = "File uploaded successfully.";
      } else {
        uploadStatusSAP.textContent = "Error uploading file.";
      }
    } catch (error) {
      console.error("Error uploading file:", error);
      uploadStatusSAP.textContent = "Error uploading file.";
    }
  } else {
    uploadStatusSAP.textContent = "Please select a file to upload.";
  }
}

async function uplaodINTEXT(e) {
  e.preventDefault(); // Prevent the default form submission behavior

  // Get the selected file from the input field
  const fileInput = document.getElementById("intextInput")
  const files = fileInput.files;

  if (files.length > 0) {
    try {


      // Append all selected files to formData
      for (let i = 0; i < files.length; i++) {
        const formData = new FormData();
        formData.append("file", files[i])
        splitMethod(formData);
      }

    } catch (error) {
      console.error("Error uploading files:", error);
      uploadStatusINTEXT.textContent = "Error uploading files.";
    }
  } else {
    uploadStatusINTEXT.textContent = "Please select a file to upload.";
  }
}

function splitMethod(formData) {
  const studyId = 1893;
  const uploadUrl = "http://localhost:9156/csrfile/Intext/upload/" + studyId;

  var headers = new Headers({
    Authorization: "Bearer " + bearerToken,
  });

  const response = fetch(uploadUrl, {
    method: "POST",
    headers: headers,
    body: formData,
  });

}

async function uplaodProtocal(e) {
  e.preventDefault(); // Prevent the default form submission behavior

  // Get the selected file from the input field
  const fileInput = document.getElementById("protocolinput")
  const file = fileInput.files[0];

  if (file) {
    try {
      const formData = new FormData();
      formData.append("file", file);

      const studyId = 1893;

      const uploadUrl = "http://localhost:9156/csrfile/Protocol/upload/" + studyId;

      var headers = new Headers({
        Authorization: "Bearer " + bearerToken,
      });

      const response = await fetch(uploadUrl, {
        method: "POST",
        headers: headers,
        body: formData,
      });

      if (response.ok) {
        uploadStatusprotocol.textContent = "File uploaded successfully.";
      } else {
        uploadStatusprotocol.textContent = "Error uploading file.";
      }
    } catch (error) {
      console.error("Error uploading file:", error);
      uploadStatusprotocol.textContent = "Error uploading file.";
    }
  } else {
    uploadStatusprotocol.textContent = "Please select a file to upload.";
  }
}
// function getPreviewURL() {
//   return new Promise((resolve, reject) => {
//     Office.onReady(function (info) {
//       if (info.host === Office.HostType.Word) {
//         Office.context.document.getFilePropertiesAsync(function (asyncResult) {
//           if (asyncResult.status === Office.AsyncResultStatus.Succeeded) {
//             // Fetch the content with authentication
//             fetch("http://localhost:8094/pls/getPlsSrcFileContent/1790", {
//               method: 'GET',
//               headers: {
//                 Authorization: "Bearer " + bearerToken,
//               }
//             })
//             .then(response => {
//               if (!response.ok) {
//                 throw new Error('Network response was not ok');
//               }
//               console.log(response.blob);
              
//               return response.blob();
//             })
//             .then(blob => {
//               const url = URL.createObjectURL(blob);
//               console.log(url);
              
//               resolve(url);
//             })
//             .catch(error => {
//               console.error("Error retrieving document content:", error);
//               reject(error);
//             });
//           } else {
//             console.error("Error retrieving document properties:", asyncResult.error.message);
//             reject(new Error(asyncResult.error.message));
//           }
//         });
//       } else {
//         reject(new Error("Not in Word context"));
//       }
//     });
//   });
// }

// // Usage
// getPreviewURL()
//   .then(url => {
//     const previewUrl = "https://docs.google.com/gview?url=https://localhost:3000/ac4ef0d2-3fce-4723-882f-c7fb57d45353&embedded=true";

    
//     document.getElementById("documentFrame").src = url;
//   })
//   .catch(error => {
//     console.error("Failed to get preview URL:", error);
//   });


function saveDocumentsss() {
  console.log(Office.context.document.url);
  Office.context.document.getFileAsync(Office.FileType.Compressed, { sliceSize: 4194304 }, function (result) {
    console.log(Office.FileType.Compressed);
    if (result.status === Office.AsyncResultStatus.Succeeded) {
      var file = result.value;
      var slices = [];
      var sliceCount = file.sliceCount;

      // eslint-disable-next-line no-inner-declarations
      function getSlice(index) {
        console.log("Getting slice " + (index + 1) + " of " + sliceCount);
        file.getSliceAsync(index, function (sliceResult) {
          if (sliceResult.status === Office.AsyncResultStatus.Succeeded) {
            slices.push(sliceResult.value.data);
            if (index + 1 < sliceCount) {
              getSlice(index + 1);
            } else {
              console.log("All slices retrieved, creating the file...");
              file.closeAsync();

              // Create a Uint8Array from all slices
              var fullContent = new Uint8Array(slices.reduce((acc, slice) => acc.concat(Array.from(slice)), []));

              var blob = new Blob([fullContent], {
                type: "application/vnd.openxmlformats-officedocument.wordprocessingml.document",
              });
              var formData = new FormData();
              formData.append("file", blob, "document.docx");

              return fetch("http://localhost:8094/pls/updateDoc/1790", {
                method: "PUT",
                headers: {
                  Authorization: "Bearer " + bearerToken,
                },
                body: formData,
              })
                .then((response) => response.json())
                .then((data) => {
                  console.log("Upload successful:", data);
                })
                .catch((error) => {
                  console.error("Upload error:", error);
                });
            }
          } else {
            console.error("Error getting slice: " + sliceResult.error.message);
          }
        });
      }

      getSlice(0);
    } else {
      console.error("Error getting file: " + result.error.message);
    }
  });
}

export function getFileContents(event) {
  const file = event.target.files[0];
  const reader = new FileReader();
  reader.onload = (_event) => {
    try {
      // Get the Base64 string without the data URL prefix
      const base64String = reader.result.split(",")[1];
      template = base64String;

      // Import the template into the document.
      importTemplate();

      // Show the Update section.
      document.getElementById("imported-section").style.display = "block";
    } catch (error) {
      console.error("Error processing file:", error);
      document.getElementById("error-message").textContent = "Failed to process file. Please try again.";
    }
  };

  reader.onerror = (error) => {
    console.error("Error reading file:", error);
    document.getElementById("error-message").textContent = "Failed to read file. Please try again.";
  };

  // Read the file as a data URL
  reader.readAsDataURL(file);
}

// Imports the template into this document.
async function importTemplate() {
  try {
    await Word.run(async (context) => {
      // Use the Base64-encoded string representation of the selected .docx file.
      context.document.insertFileFromBase64(template, "Replace", {
        importTheme: true,
        importStyles: true,
        importParagraphSpacing: true,
        importPageColor: true,
        importDifferentOddEvenPages: true,
      });
      await context.sync();
    });
  } catch (error) {
    // console.error("Error importing template:", error);
    // You can add user-friendly error handling here, such as displaying an error message
    // document.getElementById("error-message").textContent = "Failed to import template. Please try again.";
  }
}

// Helper function to convert ArrayBuffer to Base64
function arrayBufferToBase64(buffer) {
  let binary = "";
  const bytes = new Uint8Array(buffer);
  const len = bytes.byteLength;
  for (let i = 0; i < len; i++) {
    binary += String.fromCharCode(bytes[i]);
  }
  return window.btoa(binary);
}
async function getFontStyle() {
  try {
    await Word.run(async (context) => {
      const selection = context.document.getSelection();
      const font = selection.font;

      // Load the font properties
      font.load("name, size, bold, italic, underline");

      await context.sync();

      // Get the font style details
      const fontName = font.name;
      const fontSize = font.size;
      const isBold = font.bold;
      const isItalic = font.italic;
      const isUnderlined = font.underline !== "None";
      // const fontColor = font.color;

      // Display the font style details
      document.getElementById("font-style-output").textContent = `
        Font Name: ${fontName},
        Font Size: ${fontSize},
        Bold: ${isBold},
        Italic: ${isItalic},
        Underlined: ${isUnderlined}
      `;
    });
  } catch (error) {
    // console.error("Error getting font style:", error);
    // document.getElementById("font-style-output").textContent = "Failed to get font style. Please try again.";
  }
}
// function getTeseData() {
//   try {
//     Word.run(async (context) => {
//       const selection = context.document.getSelection();
//       const ooxml = selection.getOoxml();
      
//       await context.sync();

//       const selectedTextOOXML = ooxml.value;
//       const formattedText = convertOOXMLToHTML(selectedTextOOXML);
//       console.log("Formatted Text:", formattedText);

//       const payload = {
//         script: formattedText
//       };
//       const res=await fetch("http://localhost:9156/csr/getText", {
//         method: "POST",
//         headers: {
//           "Content-Type": "application/json",
//           Authorization: "Bearer " + bearerToken,
//         },
//         body: JSON.stringify(payload),
//       });
//       const data = await res.json();
      
//       const newText = data.script; 
//       console.log(selection.insertHtml(newText, "Replace"));
//       selection.insertHtml(newText, "Replace");

//       await context.sync();
      
//       console.log("Text replaced with:", newText);
      
//     });
//   } catch (error) {
//     console.error("Error getting formatted text:", error);
//     document.getElementById("formatted-output").textContent = "Select some text first!";
//   }
// }

// function convertOOXMLToHTML(ooxml) {
//   const parser = new DOMParser();
//   const xmlDoc = parser.parseFromString(ooxml, "application/xml");
//   let htmlContent = "";
  
//   // Define the namespace
//   const ns = "http://schemas.openxmlformats.org/wordprocessingml/2006/main";
  
//   const paragraphElements = xmlDoc.getElementsByTagNameNS(ns, "p");
  
//   for (let p = 0; p < paragraphElements.length; p++) {
  
//     const runElements = paragraphElements[p].getElementsByTagNameNS(ns, "r");
    
//     for (let r = 0; r < runElements.length; r++) {
//       const runElement = runElements[r];
//       const textElements = runElement.getElementsByTagNameNS(ns, "t");
//       const rPrElement = runElement.getElementsByTagNameNS(ns, "rPr")[0];
      
//       let textContent = "";
//       for (let t = 0; t < textElements.length; t++) {
//         textContent += textElements[t].textContent;
//       }
      
//       let isSubscript = false;
//       let isSuperscript = false;
//       let isItalic = false;
//       let isBold = false;

//       if (rPrElement) {
//         const vertAlignElement = rPrElement.getElementsByTagNameNS(ns, "vertAlign")[0];
//         const italicElement = rPrElement.getElementsByTagNameNS(ns, "i")[0];
//         const boldElement = rPrElement.getElementsByTagNameNS(ns, "b")[0];

//         if (vertAlignElement) {
//           const alignValue = vertAlignElement.getAttribute("w:val");
//           isSubscript = alignValue === "subscript";
//           isSuperscript = alignValue === "superscript";
//         }
        
//         isItalic = italicElement && italicElement.getAttribute("w:val") !== "false";
//         isBold = boldElement && boldElement.getAttribute("w:val") !== "false";
//       }
      
//       let formattedText = textContent;
//       if (isSubscript) formattedText = `<sub>${formattedText}</sub>`;
//       if (isSuperscript) formattedText = `<sup>${formattedText}</sup>`;
//       if (isItalic) formattedText = `<i>${formattedText}</i>`;
//       if (isBold) formattedText = `<strong>${formattedText}</strong>`;
      
//       htmlContent += formattedText;
//     }
//   }
  
//   return htmlContent;
// }



function getXML() {
  Word.run(async (context) => {
    try {
      const selection = context.document.getSelection();
      selection.load("text");
      await context.sync();

      // Check if there's any selected text
      if (selection.text.trim().length === 0) {
        throw new Error("No text selected. Please select some text and try again.");
      }

      const ooxml = selection.getOoxml();
      await context.sync();

      const selectedTextOOXML = ooxml.value;
      const bodyContent = extractBodyContent(selectedTextOOXML);
      console.log("Extracted body content:", bodyContent);

      document.getElementById("xml-output").textContent = bodyContent;
    } catch (error) {
      console.error("Error in Word.run:", error);
      let errorMessage = "An error occurred while processing the selected text.";
      
      if (error instanceof OfficeExtension.Error) {
        errorMessage += ` Office.js Error: ${error.code}, ${error.message}`;
      } else {
        errorMessage += ` ${error.message}`;
      }

      document.getElementById("xml-output").textContent = errorMessage;
    }
  });
}
function extractBodyContent(xml) {
  const parser = new DOMParser();
  const xmlDoc = parser.parseFromString(xml, "application/xml");
  const bodyElement = xmlDoc.getElementsByTagNameNS(
    "http://schemas.openxmlformats.org/wordprocessingml/2006/main",
    "body"
  )[0];
  console.log(bodyElement);
  if (bodyElement) {
    return new XMLSerializer().serializeToString(bodyElement);
  } else {
    return "No body content found";
  }
}
async function removeDoc() {
  await Word.run(async (context) => {
    // Clear the body content
    const body = context.document.body;
    body.clear();

    const numer = context.document.numer;
    await context.sync();
    // Clear the header content
    const sections = context.document.sections;
    sections.load("items");
    await context.sync();

    sections.items.forEach((section) => {
      section.getHeader("primary").clear();
      section.getHeader("firstPage").clear();
      section.getHeader("evenPages").clear();
      section.getFooter("primary").clear();
      section.getFooter("firstPage").clear();
      section.getFooter("evenPages").clear();
    });

    await context.sync();
  });
}

// New function to load document from the API
async function loadFromApi() {
  try {
    const response = await fetch("https://dev.zyliq.com/pls-service/auth/pls/generateSummaryStudyId/25/docx");
    if (!response.ok) {
      throw new Error("Network response was not ok");
    }
    const arrayBuffer = await response.arrayBuffer();
    const base64String = arrayBufferToBase64(arrayBuffer);
    template = base64String;
    importTemplate();
    // document.getElementById("imported-section").style.diysplay = "block";
  } catch (error) {
    // console.error("Error fetching document:", error);
    // document.getElementById("error-message").textContent = "Failed to load document from API. Please try again.";
  }
}

async function getFileName() {
  var intextFileCount = 0;
  try {
    const url = "http://localhost:9156/csrfile/1893";

    const headers = new Headers({
      Authorization: "Bearer " + bearerToken,
    });

    const response = await fetch(url, { headers: headers });

    if (!response.ok) {
      throw new Error(`HTTP error! Status: ${response.status}`);
    }

    const data = await response.json();

    // Assuming data is an array of CSRSourceFilesDTO objects
    data.forEach(file => {
      if (file.type === "Protocol") {
        document.getElementById('protocolFileName').innerText = file.fileName;
      } else if (file.type === "SAP") {
        document.getElementById('sapFileName').innerText = file.fileName;
      }
      else if (file.type === "Intext") {
        intextFileCount++;
        // let fileSize = file.size; // Assuming `size` is in bytes
        // console.log(`File size: ${fileSize} bytes`);
        document.getElementById('intextFileName').innerText = intextFileCount + '' + 'Intext files.';
      }
      // console.log('intext', intextFileCount);
    });

    // If you want to return the filenames
    return data.map(file => file.fileName);

  } catch (error) {
    console.error('Error fetching the filename:', error);
  }
}

