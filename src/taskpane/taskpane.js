/*
 * Copyright (c) Microsoft Corporation. All rights reserved. Licensed under the MIT license.
 * See LICENSE in the project root for license information.
 */

/* global console, document, Excel, Office */

Office.onReady((info) => {
  if (info.host === Office.HostType.Excel) {
    document.getElementById("sideload-msg").style.display = "none";
    document.getElementById("app-body").style.display = "flex";
    document.getElementById("run").onclick = () => run();
    document.getElementById("fill").onclick = () => fillFromSelection();
    document.getElementById("delete-output").onclick = () => deleteOutput();
    document.getElementById("select-words").onclick = () => load("words-start");
    document.getElementById("select-texts").onclick = () => load("texts-start");
    document.getElementById("select-results").onclick = () => load("results-start");
    document.getElementById("select-delimiter").onclick = () => load("delimiter-cell");


  }
});

let delimiter="", wordArray=[], textArray=[], wordCount, textCount, resultAddress;


async function fillFromSelection(){
  try {
    await Excel.run(async (context) => {
      const range = context.workbook.getSelectedRange();
      range.load("values");
      await context.sync();

      document.getElementById("words-start").value=range.values[0][0];
      document.getElementById("texts-start").value=range.values[1][0];
      document.getElementById("delimiter-cell").value=range.values[2][0];
      document.getElementById("results-start").value=range.values[3][0];

    });
  } catch (error) {
    console.error(error);
  }
}

async function load(targetId){
  try {
    await Excel.run(async (context) => {
      let range = context.workbook.getSelectedRange();
      
      range.load("address");
      await context.sync();
      document.getElementById(targetId).value=range.address;
      
    });
  } catch (error) {
    console.error(error);
  }
}


function getRange(context, address){
  const parts = address.split("!");
  const worksheet = parts[0];
  const cells = parts[1];
  return context.workbook.worksheets.getItem(worksheet).getRange(cells);
}

function getRangeById(context, id){
  const value = document.getElementById(id).value;
  return getRange(context, value);
}



async function processInputs(context){
  // WORDS
  {
    let wordRange = getRangeById(context,"words-start");
        
    // Read the range address
    wordRange.load("rowIndex");
    await context.sync();
    
    const top = wordRange.getOffsetRange(-wordRange.rowIndex,0);
    //const top = context.workbook.worksheets.getActiveWorksheet().getCell(0,wordRange.columnIndex);

    top.load("values");
    await context.sync();

    wordCount = top.values[0][0];
    document.getElementById("word-count").innerText = wordCount + " db";

    wordRange = wordRange.getResizedRange(wordCount-1,0);
    //wordRange.select();
    
    wordRange.load("values");
    await context.sync();
    wordArray=[];
    for(const word of wordRange.values){
      wordArray.push(word[0].toString().toLowerCase());
    }
  }  
  //TEXTS
    {
    let textRange = getRangeById(context, "texts-start");

    // Read the range address
    textRange.load("rowIndex");
    await context.sync();
    const top = textRange.getOffsetRange(-textRange.rowIndex,0);
    //const top = context.workbook.worksheets.getActiveWorksheet().getCell(0,textRange.columnIndex);

    top.load("values");
    await context.sync();

    textCount = top.values[0][0];
    document.getElementById("text-count").innerText = textCount + " db";

    textRange = textRange.getResizedRange(textCount-1,0);
    //textRange.select();

    textRange.load("values");

    await context.sync();
    textArray=[];
    for(const text of textRange.values){
      textArray.push(text[0].toString().toLowerCase());
    }
  }
  // RESULTS

  resultAddress = document.getElementById("results-start").value;

  // DELIMITER

  {
    var delimiterRange = getRangeById(context, "delimiter-cell");
    //let delimiterRange = context.workbook.worksheets.getActiveWorksheet().getRange(document.getElementById("delimiter-cell").value.toString());
    //let delimiterRange=context.workbook.worksheets.getItem("Sheet1").getRange("G4:G4");
    delimiterRange.load("values");
    await context.sync();

    
    delimiter=delimiterRange.values[0][0];
    document.getElementById("delimiter").innerText = "\""+delimiter+"\"";
  }
}

async function deleteOutput() {
  try {
    await Excel.run(async (context) => {
      const outputStart = getRangeById(context, "results-start");
      const outputColumns = outputStart.getResizedRange(100000,16);
      outputColumns.clear("Contents");
    });
  } catch (error) {
    console.error(error);
    document.getElementById("error").innerText="(Hiba: "+error+")";
  }
}
async function run() {
  try {
    await Excel.run(async (context) => {
      document.getElementById("error").innerText="";
      const runButton = document.getElementById("run-text");
      runButton.innerText="Futtatás...";
      await processInputs(context);
      const rootCell = getRange(context, resultAddress);
      let activeCell = rootCell;

      runButton.innerText="Futtatás... (1)";
      await context.sync();
      await runSingleCounter(activeCell, false);
      activeCell=rootCell.getOffsetRange(0,3);
      await context.sync();

      runButton.innerText="Futtatás... (2)";
      await context.sync();
      await runSingleCounter(activeCell, true);
      activeCell=rootCell.getOffsetRange(0,6);
      await context.sync();

      runButton.innerText="Futtatás... (3)";
      await runPairsCounter(activeCell);
      await context.sync();
      activeCell=rootCell.getOffsetRange(0,10);
      runButton.innerText="Futtatás... (4)";
      
      await runCountTags(activeCell);
      await context.sync();
      activeCell=rootCell.getOffsetRange(0,12);
      runButton.innerText="Futtatás... (5)";
      
      await runCreateHistogram(activeCell);
      await context.sync();
      activeCell=rootCell.getOffsetRange(0,15);
      runButton.innerText="Futtatás... (6)";
      
      await runCreateMatchHistogram(activeCell);
      await context.sync();

    });
  } catch (error) {
    console.error(error);
    document.getElementById("error").innerText="(Hiba: "+error+")";

  }finally{
    document.getElementById("run-text").innerText="Futtatás";

  }
}

function refreshElement(element){
  element.style.display="none";
  element.style.display="block";
}


async function runSingleCounter(resultRange, ordered) {
      let occurences = new Map();
      let total = 0;
      for(const word of wordArray){
        let occurence = 0;
        for(const text of textArray){
          //occurence += text[0].split(word[0]).length - 1;
          //occurence += (text[0].match(new RegExp(word[0],"gi")) || []).length;
          //occurence += (text[0].match(new RegExp(word[0],"gi")) || []).length>0?1:0;
          occurence += text.includes(word);
          continue;
        }
        occurences.set(word,occurence);
        if(word != "")
          total +=occurence;
      }
      console.log(occurences);
      console.log("Total:"+total);

      if(ordered){
        occurences = new Map([...occurences.entries()].sort((a, b) => b[1] - a[1]));
        resultRange=resultRange.getResizedRange(occurences.size-1,2);
      }else{
        resultRange=resultRange.getResizedRange(occurences.size-1,1);
      }

      if(ordered){
        occurences = new Map([...occurences.entries()].sort((a, b) => b[1] - a[1]));
      }

      //let resultRange = context.workbook.worksheets.getActiveWorksheet().getRange(resultAddress);
      //resultRange=resultRange.getResizedRange(occurences.size-1,1);
      //resultRange.select();
      resultRange.values=[];
      let sum=0;
      for(const [word,count] of occurences){
        if(ordered){
          if(word == ""){
            resultRange.values.push([word,count,""]);
          }else{
            sum += count;
            let relativesum = sum/total;
            resultRange.values.push([word,count,relativesum]);
          }
        }else{
          resultRange.values.push([word,count]);
        }
      }
      //await context.sync();

}


async function runPairsCounter(resultRange) {
      const occurences = {};
      for(i=0;i<wordArray.length-1;i++){
        for(j=i+1;j<wordArray.length;j++){
          for(const text of textArray){
            if(text.includes(wordArray[i]) && text.includes(wordArray[j])){
              if(occurences[i+"x"+j]==undefined){
                occurences[i+"x"+j]=0;
              }
              occurences[i+"x"+j]+=1;
            }
          }
        }
      }

      console.log(occurences);

      //let resultRange = context.workbook.worksheets.getActiveWorksheet().getRange(resultAddress);
      resultRange=resultRange.getResizedRange(Object.keys(occurences).length-1,2);
      //resultRange.select();
      resultRange.values=[];
      for(const key in occurences){
        const keys = key.split("x");
        resultRange.values.push([parseInt(keys[0])+1,parseInt(keys[1])+1,occurences[key]]);
      }

}

String.prototype.count = function(search) {
  var m = this.match(new RegExp(search.toString().replace(/(?=[.\\+*?[^\]$(){}\|])/g, "\\"), "gi"));
  return m ? m.length:0;
}

async function runCountTags(resultRange) {
      const occurences = [];
      for(const text of textArray){
        const array = text.split(delimiter);
        if(array.length==1 && array[0]==""){
          occurences.push(0);
        }else{
          const occurence = text.split(delimiter).length;
          occurences.push(occurence);
        }
      }

      console.log(occurences);

      //let resultRange = context.workbook.worksheets.getActiveWorksheet().getRange(resultAddress);
      resultRange=resultRange.getResizedRange(occurences.length-1,0);
      //resultRange.select();
      resultRange.values=[];
      for(const occurence of occurences){
        resultRange.values.push([occurence]);
      }

}


async function runCreateHistogram(resultRange) {
      let maxTagCount=0;
      const tagCounts = {};
      for(const text of textArray){
        const array = text.split(delimiter);
        let tagCount=0;
        if(array.length !=1 || array[0] != ""){
          tagCount = text.split(delimiter).length;
        }
        if(tagCount>maxTagCount){
          maxTagCount=tagCount;
        }
        if(tagCounts[tagCount+""] == undefined){
          tagCounts[tagCount+""] = 0;
        }
        tagCounts[tagCount+""] += 1;
        
      }

      console.log(tagCounts);

      //let resultRange = context.workbook.worksheets.getActiveWorksheet().getRange(resultAddress);
      resultRange=resultRange.getResizedRange(maxTagCount,1);
      //resultRange.select();
      resultRange.values=[];
      for(i=0;i<=maxTagCount;i++){
        resultRange.values.push([i,tagCounts[i+""]==undefined?0:tagCounts[i+""]]);
      }

}


async function runCreateMatchHistogram(resultRange) {

        const occurences = [];
        for(const text of textArray){
          let occurence = 0;
          for(const word of wordArray){
            occurence += text.count(word);//(text.match(new RegExp(word,"gi")) || []).length;
          }
          occurences.push(occurence);
          //occurences.push(occurence);
        }

        console.log(occurences);

        //let resultRange = context.workbook.worksheets.getActiveWorksheet().getRange(resultAddress);
        resultRange=resultRange.getResizedRange(occurences.length,1);
        //resultRange.select();
        resultRange.values=[];
        for(i=0;i<=occurences.length;i++){
          resultRange.values.push([i+1,occurences[i]==undefined?0:occurences[i]]);
        }
        /*for(const occurence in occurences){
          resultRange.values.push([occurence,occurences[occurence]]);
        }*/


}



