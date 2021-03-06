const R=RAMDA.R
Number.prototype.to26 = function (suffix) {
    suffix = String.fromCharCode((this % 26) + 65) + (suffix || '');
    return this >= 26 ? (Math.floor(this / 26) - 1).to26(suffix) : suffix;
};
//const daggy=DAGGY.daggy
(function (global, factory) {
  typeof exports === 'object' && typeof module !== 'undefined' ? factory(exports) :
  typeof define === 'function' && define.amd ? define(['exports'], factory) :
  (factory((global.P = {})));
}(this, (function (exports) {'use strict';
const version=()=>'v6';

const getTopCell= function({ss,name='Parametrs'}) {
const topCell=getNameRange_({ name: name, ss:ss })
return topCell
}
const getRanges=({topCell,isHeader=true,isTotal=false,isFilter=true})=>{
const headerStartRow=isHeader?-1:0
const lastRowRange=getLastRowByCell_(topCell);
const lastColumnRange=lastNoEmtyCol_(topCell);
const dataRangeEndRow=isHeader?lastRowRange-1:lastRowRange;
const dataRange=topCell.offset(headerStartRow, 0,dataRangeEndRow,lastColumnRange);
const headersRange=isHeader?topCell.offset(headerStartRow, 0,1,lastColumnRange):undefined;
const totalsRange=isTotal?topCell.offset(lastRowRange, 0,1,lastColumnRange):undefined;
return {
topCell:topCell,
//lastRowRange:lastRowRange,
//lastColumnRange:lastColumnRange,
dataRange:dataRange,
headersRange:headersRange,
totalsRange:totalsRange,
}
}
 /**
     * Получим диапазон в активной книге 
     * по имени именнованованого диапазона
     *
     * @param {*} { name } 
     * @returns {Range}
     */
    function getNameRange_({ name,ss }) {
    
        return ss.getRangeByName(name)

    }
//Последняя непустая ячейка в 1 столбце
function getLastRowByColumn_(range){
  while(range.length>0 && range[range.length-1][0]=='') range.pop();
  return range.length;
}
      
 const toObjectAndHeaders = (array,header,A1Not) =>{return {arrObj:toObject(array) , headers:header,A1Not:A1Not}}

 //Последняя непустая ячейка в столбце topCell относительно ячейки topCell
 const getLastRowByCell_=(topCell)=>getLastRowByColumn_(topCell.offset(0, 0, topCell.getSheet().getLastRow()).getValues()) 
 const isEmty=(x)=>x===''
 //Последняя непустая ячейка в Строке НАД topCell относительно ячейки topCell
 const getLastColumnByCell_=(topCell)=>R.findIndex(isEmty,topCell.offset(-1, 0,1, topCell.getSheet().getLastColumn()).getValues())-1//getLastRowByColumn_(transpose_(topCell.offset(-1, 0,1, topCell.getSheet().getLastColumn()).getValues())) 

 //Последняя непустая ячейка в Строке НАД topCell относительно ячейки topCell
 const lastNoEmtyCol_=(topCell)=>{
 var range=topCell.offset(-1, 0,1, topCell.getSheet().getLastColumn())
 const val =range.getValues()[0]
 const ind=R.findIndex(isEmty,val)
 return ind
 };
 
//var g={name:"j"}
//g.prototype.getValues=(name)=>g.name
const getValues=(ranges)=>{
var keys=Object.keys(ranges)
var v={}
keys.forEach(function(key) {
console.log(ranges[key])
if(ranges[key]){
 v[key]= ranges[key].getValues();
 }  
});

return v
}

const getA1Nots=(ranges)=>{
var keys=Object.keys(ranges)
var v={}
keys.forEach(function(key) {
console.log(ranges[key])
if(ranges[key]){
 v[key]= ranges[key].getA1Notation();
 }  
});
return v
}
const getObjects=(topCell,values,A1Not)=>{
//const v=values({topCell,isHeader=true,isTotal=false,isFilter=true})
return {
objDataAndHeaders:toObjectAndHeaders(values.dataRange,values.headersRange[0],A1Not),
hashTable:getParametrsByNameRange_(topCell),
dataValues:values.dataValues,
}
}

const CONFIG= function({ssId,name='Parametrs',isHeader=true,isTotal=false,isFilter=true}) {
  
// console.log(ssId)
const ss=SpreadsheetApp.openById(ssId)
if(!ss){ 
  return {error:`Таблица с id ${ssId} не найдена `,result:{}}
}
const topCell=getTopCell({ name: name, ss:ss })
if (!topCell) {
  return {error:` Именннованный диапазон с именем ${name} в таблице с id ${ssId} не найден `,result:{}}
}

const ranges=getRanges(({topCell:topCell,isHeader:isHeader,isTotal:isTotal,isFilter:isFilter}))
const values=getValues(ranges)
return getObjects(values)
}

//}
  /**
     * Транспортирование масива
     * 
     * @param {*} array
     */
 const transpose_ = array => array.reduce((r, a) => a.map((v, i) => [...(r[i] || []), v]), []);  
/**
 * Превратим таблицу в обьект [{},{}] заголовки станут ключами
 *
 * @param {Array} array
 * @returns

 */ 
 const toObject = (array) => { const keys = array.shift();
                              return array.map((values) => { return keys.reduce((o, k, i) => { o[k] = values[i]; return o }, {}) })
};

 function isJSON(MyTestStr){
    try {
        var MyJSON = JSON.stringify(MyTestStr);
        var json = JSON.parse(MyJSON);
        if(typeof(MyTestStr) == 'string')
            if(MyTestStr.length == 0)
                return false;
    }
    catch(e){
        return false;
    }
    return true;
}
const getListObjects=(ssId)=>{
const ss=SpreadsheetApp.openById(ssId)
var result = getNamedRanges(ssId);
return Object.keys(result).filter(el=>{
var note=ss.getRangeByName(el).getNote()
return isJSON(note)?JSON.parse(note).hasOwnProperty("description"):false;
}).map(el=>{
const topCell=getTopCell({ name: result[el], ss:ss });
const sheeetName=topCell.getSheet().getName();
return [el,sheeetName,getRanges({topCell:topCell})['dataRange'].getA1Notation()]});
 
}

const getParametrsByNameRange_=(topCell)=>toObject( transpose_(getValues_(getRangeTopCellTwoColumns_(topCell))))

//Функции для импорта
const getColumnLetterFromHeader_ =R.curry((objHeader,columnNameResponse)=>objHeader.indexOf(columnNameResponse).to26())
const getColumnIndexFromHeader_ =R.curry((objHeader,columnNameResponse)=>objHeader.indexOf(columnNameResponse))
const getValueByKeyName_=R.curry((obj,paramName) => R.view(R.lensProp(paramName),obj))
const getColumnIndexTopCell_=R.curry((topCell,objHeader,columnNameResponse)=>topCell.getColumn()+objHeader.indexOf(columnNameResponse)) 
const getColumnLetterTopCell_=R.curry((topCell,objHeader,columnNameResponse)=>(topCell.getColumn()+objHeader.indexOf(columnNameResponse)).to26())
const runMetod_=(metod,obj)=>(obj)=>obj[metod]()
const getValues_=runMetod_('getValues')
const match = R.curry((what, s) => s.match(what));
   //Получим диапазон расширенный  относительно ячейки topCell на column столбцов
 const getRangeTopCell_=(column,topCell)=>(topCell)=>topCell.offset(0, 0,getLastRowByCell_(topCell),column)
 //Получим диапазон расширенный на 1 столбец относительно ячейки topCell 
 const getRangeTopCellTwoColumns_=getRangeTopCell_(2) 
//Вы можете использовать это для расчета пересечения двух диапазонов. 
//Требуется объект в виде:{rg1:'A1Notation String',rg2:'A1Notation String'}

function calculateIntersection(rgObj) {
  var iObj={};
  var ss=SpreadsheetApp.getActive();
  var sh=ss.getActiveSheet();
  var rg1=sh.getRange(rgObj.rg1);
  var rg2=sh.getRange(rgObj.rg2);
  var iObj={rg1colst:rg1.getColumn(),rg1colen:rg1.getColumn()+rg1.getWidth()-1,rg1rowst:rg1.getRow(),rg1rowen:rg1.getRow()+rg1.getHeight()-1,rg2colst:rg2.getColumn(),rg2colen:rg2.getColumn()+rg2.getWidth()-1,rg2rowst:rg2.getRow(),rg2rowen:rg2.getRow()+rg2.getHeight()-1};
  if(iObj.rg1colst>iObj.rg2colen || iObj.rg1colen<iObj.rg2colst || iObj.rg1rowst>iObj.rg2rowen || iObj.rg1rowen<iObj.rg2rowst || iObj.rg2colst>iObj.rg1colen || iObj.rg2colen<iObj.rg1colst || iObj.rg2rowst>iObj.rg1rowen || iObj.rg2rowen<iObj.rg1rowst) {
    return ;
  }else{
    var vA1=rg1.getValues();
    var v1=[];
    var vA2=rg2.getValues();
    var v2=[];
    for(var i=0;i<vA1.length;i++){
      for(var j=0;j<vA1[i].length;j++){
        var s=Utilities.formatString('(%s,%s)', iObj.rg1rowst+i,iObj.rg1colst+j);
        v1.push(s);
      }
    }
    for(var i=0;i<vA2.length;i++){
      for(var j=0;j<vA2[i].length;j++){
        var s=Utilities.formatString('(%s,%s)', iObj.rg2rowst+i,iObj.rg2colst+j);
        v2.push(s);
      }
    }
    var oA=[];
    for(var i=0;i<v1.length;i++){
      var idx=v2.indexOf(v1[i]);
      if(idx>-1){
        oA.push(v2[idx]);
      }
    }
    return oA//Utilities.formatString('Intersecting Cells: %s', oA.join(', '));
  }
} 

//https://gist.github.com/tanaikech/aa744c9a15818c002d90eaea6b4efd03
function getNamedRanges(spreadsheetId) {
    var ss = SpreadsheetApp.openById(spreadsheetId);
    var sheetIdToName = {};
    ss.getSheets().forEach(function(e) {
        sheetIdToName[e.getSheetId()] = e.getSheetName();
    });
    var result = {};
    Sheets.Spreadsheets.get(spreadsheetId, {fields: "namedRanges"})
        .namedRanges.forEach(function(e) {
            var sheetName = sheetIdToName[e.range.sheetId.toString()];
            var a1notation = ss.getSheetByName(sheetName).getRange(
                e.range.startRowIndex + 1,
                e.range.startColumnIndex + 1,
                e.range.endRowIndex - e.range.startRowIndex,
                e.range.endColumnIndex - e.range.startColumnIndex
            ).getA1Notation();
            result[e.name] = sheetName + "!" + a1notation;
        });
    return result;
}
function mainGetNamedRanges() {
    var spreadsheetId = "### spreadsheet ID ###";
    var result = getNamedRanges(spreadsheetId);
    Logger.log(JSON.stringify(result));
}
 const getIntersection=match(/\((.*)\)/) 
 const NoLabel=R.complement(match(/.*_label/))
exports.R=R;
exports.match=match;
exports.NoLabel=NoLabel
exports.getIntersection=getIntersection
exports.getA1Nots=getA1Nots;
exports.toObject=toObject;
exports.getTopCell=getTopCell;
exports.getRanges=getRanges;
exports.getValues=getValues;
exports.getObjects=getObjects;
exports.getColumnLetterFromHeader=getColumnLetterFromHeader_;
exports.getValueByKeyName=getValueByKeyName_;
exports.getColumnIndexFromHeader=getColumnIndexFromHeader_;
exports.getColumnLetterTopCell=getColumnLetterTopCell_;
exports.getColumnIndexTopCell=getColumnIndexTopCell_;
exports.version = version;
exports.CONFIG = CONFIG;
exports.calculateIntersection=calculateIntersection;
exports.getNamedRanges=getNamedRanges;
exports.getListObjects=getListObjects;



Object.defineProperty(exports, '__esModule', { value: true });
})));



