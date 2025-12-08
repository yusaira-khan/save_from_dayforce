function main() {
  const sheets = SpreadsheetApp.getActiveSpreadsheet().getSheets();
  const startSlice = 0
  let idx = 0
  for (let s of sheets.slice(startSlice)){
    console.log(`[${idx+startSlice}] reformatting ${s.getName()}`)

    reformatTable_(s)
    idx++;
  }
}

function removeUnwanted_(s, m=null){
  const range = s.getDataRange()
  const numRows = range.getNumRows()
  const rowsToRemove = []

  for (let offset=0; offset<numRows; offset++){
    const row = range.offset(offset, 0, 1);
    if (isUnwanted_(row)){
      rowsToRemove.push(row.getRow())
    }
  }

  if (null == m){
    removeRowsInBatches_(s, rowsToRemove)
    console.log(`removed ${rowsToRemove.length} rows from ${s.getName()}`)
  }else{
    m(s, rowsToRemove)
  }
}

function isUnwanted_(row){
    const rawValues = row.getValues()[0];
    const values = rawValues.filter(Boolean);
    if (values.length <= 1){
      return true
    }
    if (values.length > 3 && values[0] === "Earnings" && values[1] === "Begin Date"	&& values[2] === "End Date"){
      return true
    }
    if(values[0].startsWith("XXX") || values[0].startsWith("Direct Deposit")){
      return true
    }
    return false
}

function removeRowsInBatches_(sheet, rows){
  const batches = gatherRows_(sanitizeRows_(rows))
  for (const rowBatch of batches){
    sheet.deleteRows(rowBatch.start, rowBatch.howMany)
  }
}

function gatherRows_(rows){
  const batches = []
  for (const currentRow  of rows){
    let currentSet = -1
    if (batches.length>0){
      currentSet = batches[batches.length-1].start
    }
    if (currentSet-1 === currentRow){
      batches[batches.length-1].start = currentRow
      batches[batches.length-1].howMany++
    } else {
      batches.push({start:currentRow, howMany: 1})
    }
  }
  return batches
}

function removeRows1by1_(sheet, rows){
  for (const num of sanitizeRows_(rows)){
    sheet.deleteRow(num)
  }
}

function sanitizeRows_(rows){
  const unique = [...new Set(rows)]
  unique.sort().reverse() // descending order
  return unique
}

function removeDuplicates_(s){
  const range = s.getDataRange()
  const numRows = range.getNumRows()
  const rangeValues = getValues_(range)
  const rowsToRemove = []
  for (let offset=0; offset<numRows; offset++){
    const rowNum = range.offset(offset, 0, 1).getRow();

    if (isDuplicateIn_(rangeValues, offset)){
      rowsToRemove.push(rowNum)
    }
  }
  removeRowsInBatches_(s, rowsToRemove)
  console.log(`removed ${rowsToRemove} rows from ${s.getName()}`)
}

function isDuplicateIn_(allRowValues, testRowIdx){
  const testRow = allRowValues[testRowIdx]
  const duplicates = allRowValues.filter((candidate, candidateIdx)=> candidateIdx != testRowIdx && candidate[0]==testRow[0] )
  if (duplicates.length === 0){
    return false
  } else if(duplicates.length > 1) {
    throw new Error(`Unhandled! ${duplicates.length} possible duplicates found for ${testRowIdx}! underTest=${JSON.stringify(testRow)},candidates=${JSON.stringify(duplicates)} `)
  }else {
    const candidate = duplicates[0];
    if (testRow.length < candidate.length){
      return isSubrow_(candidate, testRow)
    }
    if (testRow.length === candidate.length){
      throw new Error(`Unhandled! ${testRow.length} has exact length match! underTest=${JSON.stringify(testRow)} candidate=${JSON.stringify(candidate)}`)
    }
    return false
  }
  //isDuplicateIn_([[1,2,3],[1]],1) => true
  //isDuplicateIn_([[1,2,3],[1]],0) => false
  //isDuplicateIn_([[1,2,3],[0,2,3]],0) => false
}

function isSubrow_(big, small){
  for(let idx=0; idx<small.length; idx++){
    if(big[idx]===small[idx]){
      continue;
    }else{
      return false
    }
  }
  return true

  // isSubrow_([1,2,3], [1]) => true
  // isSubrow_([1,2,3], [1,2,3]) => true
  // isSubrow_([0,1,2,3], [1]) => false
}

function getValues_(range){
  const r=range.getValues()
  return r.map((c)=>c.filter(Boolean))
}

function reformatTableSlow_(s){
  const range = s.getDataRange()
  const numRows = range.getNumRows()
  const numCols = range.getNumRows()

  for (let r=0; r<numRows; r++){
    const row = range.offset(r, 0, 1);
    handleAggregate_(row);
    for (let c=0; c<numCols; c++){
      const cell = range.offset(r, c, 1, 1)
      handleNumberFormat_(cell)
    }
    handleBadRow_(row)
  }
}

const COLOR_LIGHT_CYAN_2 = "#a2c4c9"
const COLOR_GRAY = "#cccccc"

const AGGREGATE_COLOR = COLOR_LIGHT_CYAN_2
const AGGREGATE_NAMES = new Set(["Earnings", "Taxable Benefits", "Taxes", "Net Pay", "Pre-Tax Deductions", "Post-Tax Deductions", "Reimbursements", "Memo Information"])

function handleAggregate_(range){
  if (AGGREGATE_NAMES.has(range.getValue())){
    range.setBackground(AGGREGATE_COLOR)
  }
}

const DOLLAR_FORMAT = '"$"#,##0.00;"$"\(#,##0.00\);$0.00;@'
const DECIMAL_4_REGEX = /^\d+\.\d{4}$/
const DECIMAL_4_FORMAT = "0.0000"
const DECIMAL_2_REGEX = /^\d+\.\d{2}$/
const DECIMAL_2_FORMAT = "0.00"

function handleNumberFormat_(cell){
  const value = cell.getValue()
  if (cell.isBlank() || typeof(value) === "number"){
    return;
  } else if (value.indexOf("$")!==-1){
    cell.setNumberFormat(DOLLAR_FORMAT)
  } else if (DECIMAL_4_REGEX.test(value)){
    cell.setNumberFormat(DECIMAL_4_FORMAT)
  } else if (DECIMAL_2_REGEX.test(value)){
    cell.setNumberFormat(DECIMAL_2_FORMAT)
  }
}

function handleBadRow_(row){
  const values = row.getValues()[0]
  if (values.length>=5){
    if (values[0]==='' && values[1]==='' && values[3]==='Amount' && values[4]==='' && values[5]==='Amount' ){
      row.getSheet().deleteRow(row.getRow())
    }
  }
}

function reformatTable_(s){
  const tableRange = getProperTable_(s.getDataRange())
  adjustColumnFormats_(tableRange)

  const numRows = tableRange.getNumRows()
  for (let r=0; r<numRows; r++){
    const row = tableRange.offset(r, 0, 1);

    handleAggregate_(row);
  }
}

const DOLLAR_FORMAT = '"$"#,##0.00;"$"\(#,##0.00\);$0.00'
const DECIMAL_4_FORMAT = "0.0000"
const DECIMAL_2_FORMAT = "0.00"

function adjustColumnFormats_(tableRange){
  const s= tableRange.getSheet()
  const numRows = tableRange.getNumRows()
  const numColumns= tableRange.getNumColumns();
  s.setColumnWidth(1, 150)
  s.setColumnWidths(2, numColumns-1, 100)
  tableRange.offset(0, 1, numRows, 1).setNumberFormat(DECIMAL_2_FORMAT)
  tableRange.offset(0, 2, numRows, 1).setNumberFormat(DECIMAL_4_FORMAT)
  tableRange.offset(0, 3, numRows, 1).setNumberFormat(DOLLAR_FORMAT)
  tableRange.offset(0, 4, numRows, 1).setNumberFormat(DECIMAL_2_FORMAT)
  tableRange.offset(0, 5, numRows, 1).setNumberFormat(DOLLAR_FORMAT)
}

function getProperTable_(dataRange){
  const numColumns = dataRange.getNumColumns()
  const sheet = dataRange.getSheet()

  const mainHeaderRow = findMainHeader_(dataRange)
  const lastHeaderRow = findLastHeader_(dataRange)

  const numTableRows = 1 + lastHeaderRow.getRow() - mainHeaderRow.getRow()
  if (mainHeaderRow.getRow() > 1){
    const tableRangeSpec = sheet.getRange(mainHeaderRow.getRow(), 1, numTableRows+3)
    sheet.moveRows(tableRangeSpec, 1)
  }
  const tableRange = sheet.getRange(1, 1, numTableRows, numColumns)
  return tableRange
}

const MAIN_HEADER_TEXT_OLD = "Detail at the time pay statement issued"
const MAIN_HEADER_TEXT_REPLACE = "Pay Statement"

function findMainHeader_(range){
  for (let r=0; r<range.getNumRows(); r++){
    const row = range.offset(r, 0, 1);
      if (row.getValue()===MAIN_HEADER_TEXT_OLD ){
        row.getCell(1, 1).setValue(MAIN_HEADER_TEXT_REPLACE)
        return row
      } else if (row.getValue()===MAIN_HEADER_TEXT_REPLACE){
        return row
      }
  }
  throw new Error("Unhandled!")
}

function findLastHeader_(range){
  for (let r=0; r<range.getNumRows(); r++){
    const row = range.offset(r, 0, 1);
      if (row.getValue()==="Net Pay"){
        handleBadRow_(row.offset(-1, 0))
        return row
      }
  }
  throw new Error("Unhandled!")
}

function formatMainHeader_(s){
  if (!s.getRange(1, 2).isPartOfMerge()){
    if (!s.getRange(1, 3).isBlank()){
      s.getRange(1, 5).setValue(s.getRange(1, 3).getValue())
    }
    s.getRange(1, 2, 1, 3).merge()
    SpreadsheetApp.flush()
    s.getRange(1, 5, 1, 2).merge()
    SpreadsheetApp.flush()
  }
}
