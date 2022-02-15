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

function getRangeById(context, id){
  const element = document.getElementById(id).value;
  return context.workbook.worksheets.getActiveWorksheet().getRange(element);
}

async function processInputs(context){
  // WORDS
  {
    let wordRange = getRangeById(context,"words-start");
        
    // Read the range address
    wordRange.load("columnIndex");
    await context.sync();
    
    const top = context.workbook.worksheets.getActiveWorksheet().getCell(0,wordRange.columnIndex);

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
    textRange.load("columnIndex");
    await context.sync();

    const top = context.workbook.worksheets.getActiveWorksheet().getCell(0,textRange.columnIndex);

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

async function run() {
  try {
    await Excel.run(async (context) => {
      document.getElementById("run-text").innerText="Futtatás...";
      await processInputs(context);
      const rootCell = context.workbook.worksheets.getActiveWorksheet().getRange(resultAddress);
      let activeCell = rootCell;

      await runSingleCounter(activeCell);
      activeCell=rootCell.getOffsetRange(0,2);
      await runPairsCounter(activeCell);
      activeCell=rootCell.getOffsetRange(0,6);
      await runCountTags(activeCell);
      activeCell=rootCell.getOffsetRange(0,8);
      await runCreateHistogram(activeCell);
      activeCell=rootCell.getOffsetRange(0,11);
      await runCreateMatchHistogram(activeCell);


      document.getElementById("run-text").innerText="Futtatás";
    });
  } catch (error) {
    console.error(error);
  }
}


async function runSingleCounter(resultRange) {
  try {
      const occurences = [];
      for(const word of wordArray){
        let occurence = 0;
        for(const text of textArray){
          //occurence += text[0].split(word[0]).length - 1;
          //occurence += (text[0].match(new RegExp(word[0],"gi")) || []).length;
          //occurence += (text[0].match(new RegExp(word[0],"gi")) || []).length>0?1:0;
          occurence += text.includes(word);
          continue;
        }
        occurences.push(occurence);
      }
      console.log(occurences);

      //let resultRange = context.workbook.worksheets.getActiveWorksheet().getRange(resultAddress);
      resultRange=resultRange.getResizedRange(occurences.length-1,0);
      //resultRange.select();
      resultRange.values=[];
      for(const occurence of occurences){
        resultRange.values.push([occurence]);
      }
      //await context.sync();

  } catch (error) {
    console.error(error);
  }
}


async function runPairsCounter(resultRange) {
  try {

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
      
  } catch (error) {
    console.error(error);
  }
}

String.prototype.count = function(search) {
  var m = this.match(new RegExp(search.toString().replace(/(?=[.\\+*?[^\]$(){}\|])/g, "\\"), "gi"));
  return m ? m.length:0;
}

async function runCountTags(resultRange) {
  try {


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

  } catch (error) {
    console.error(error);
  }
}


async function runCreateHistogram(resultRange) {
  try {

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

  } catch (error) {
    console.error(error);
  }
}


async function runCreateMatchHistogram(resultRange) {
  try {

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

  } catch (error) {
    console.error(error);
  }
}



