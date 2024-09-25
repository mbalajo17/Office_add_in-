Office.onReady((info) => {
    if (info.host === Office.HostType.Word) {
      document.getElementById("sideload-msg").style.display = "none";
      document.getElementById("get-textfor").addEventListener("click", getTeseData);
    }
  });
  const bearerToken =
  "eyJ1c2VyIjp7ImlkIjoxMiwidXNlck5hbWUiOiJiYWxhamltb2hhbkBzeW1iaWFuY2UuY29tIiwicm9sZXMiOnsiaWQiOjMsInJvbGVOYW1lIjoiVXNlciJ9LCJhdXRob3JpdGllcyI6bnVsbCwiZmlyc3ROYW1lIjoiYmFsYWppIiwibGFzdE5hbWUiOiJtb2hhbiIsIm1pZGRsZU5hbWUiOm51bGwsInNob3dIZWxwIjp0cnVlLCJvcmdhbml6YXRpb24iOnsiaWQiOjEsIm9yZ05hbWUiOiJTeW1iaWFuY2UifX0sImFsZyI6IkhTNTEyIn0.eyJzdWIiOiJiYWxhamltb2hhbkBzeW1iaWFuY2UuY29tIiwiYXV0aG9yaXRpZXMiOlsiVXNlciJdLCJpYXQiOjE3MjUyNzc5NzEsImV4cCI6MTcyNTM2NDM3MX0.ruxwt7i6auSwtRd5hU2651EVL3EgW7KFpfQ6j0_Wz6georf8_upwr7RDSmnJ4tt1nf-JjInCU9Ox0AQOnVRVMw";


  function getTeseData() {
    try {
        // Show loader
        document.getElementById('loader').style.display = 'block';
        document.getElementById('get-textfor').disabled = true;

        Word.run(async (context) => {
            const selection = context.document.getSelection();
            const ooxml = selection.getOoxml();
            
            await context.sync();

            const selectedTextOOXML = ooxml.value;
            const formattedText = convertOOXMLToHTML(selectedTextOOXML);
            console.log("Formatted Text:", formattedText);

            const payload = {
                script: formattedText
            };
            const res = await fetch("http://localhost:9156/csr/getText", {
                method: "POST",
                headers: {
                    "Content-Type": "application/json",
                    Authorization: "Bearer " + bearerToken,
                },
                body: JSON.stringify(payload),
            });
            const data = await res.json();
            
            const newText = data.script; 
            console.log(selection.insertHtml(newText, "Replace"));
            selection.insertHtml(newText, "Replace");

            await context.sync();
            
            console.log("Text replaced with:", newText);

        }).catch(error => {
            console.error("Error in Word.run:", error);
        }).finally(() => {
            // Hide loader and re-enable button
            document.getElementById('loader').style.display = 'none';
            document.getElementById('get-textfor').disabled = false;
        });
    } catch (error) {
        console.error("Error getting formatted text:", error);
        document.getElementById("formatted-output").textContent = "Select some text first!";
        // Hide loader and re-enable button in case of error
        document.getElementById('loader').style.display = 'none';
        document.getElementById('get-textfor').disabled = false;
    }
}
  
  function convertOOXMLToHTML(ooxml) {
    const parser = new DOMParser();
    const xmlDoc = parser.parseFromString(ooxml, "application/xml");
    let htmlContent = "";
    
    // Define the namespace
    const ns = "http://schemas.openxmlformats.org/wordprocessingml/2006/main";
    
    const paragraphElements = xmlDoc.getElementsByTagNameNS(ns, "p");
    
    for (let p = 0; p < paragraphElements.length; p++) {
    
      const runElements = paragraphElements[p].getElementsByTagNameNS(ns, "r");
      
      for (let r = 0; r < runElements.length; r++) {
        const runElement = runElements[r];
        const textElements = runElement.getElementsByTagNameNS(ns, "t");
        const rPrElement = runElement.getElementsByTagNameNS(ns, "rPr")[0];
        
        let textContent = "";
        for (let t = 0; t < textElements.length; t++) {
          textContent += textElements[t].textContent;
        }
        
        let isSubscript = false;
        let isSuperscript = false;
        let isItalic = false;
        let isBold = false;
  
        if (rPrElement) {
          const vertAlignElement = rPrElement.getElementsByTagNameNS(ns, "vertAlign")[0];
          const italicElement = rPrElement.getElementsByTagNameNS(ns, "i")[0];
          const boldElement = rPrElement.getElementsByTagNameNS(ns, "b")[0];
  
          if (vertAlignElement) {
            const alignValue = vertAlignElement.getAttribute("w:val");
            isSubscript = alignValue === "subscript";
            isSuperscript = alignValue === "superscript";
          }
          
          isItalic = italicElement && italicElement.getAttribute("w:val") !== "false";
          isBold = boldElement && boldElement.getAttribute("w:val") !== "false";
        }
        
        let formattedText = textContent;
        if (isSubscript) formattedText = `<sub>${formattedText}</sub>`;
        if (isSuperscript) formattedText = `<sup>${formattedText}</sup>`;
        if (isItalic) formattedText = `<i>${formattedText}</i>`;
        if (isBold) formattedText = `<strong>${formattedText}</strong>`;
        
        htmlContent += formattedText;
      }
    }
    
    return htmlContent;
  }
  