function main() {
  const sheets = SpreadsheetApp.getActiveSpreadsheet().getSheets();
  for (let s of sheets.slice(55+30)){
    row_(s)
  }
}

function row_(s, m=null){
  const range = s.getDataRange()
  const num_rows = range.getNumRows()
  const rows_to_remove = []

  for (let offset=0; offset<num_rows; offset++){
    const row = range.offset(offset, 0, 1);
    if (is_unwanted_(row)){
      rows_to_remove.unshift(row.getRow())
    }
  }

  if (null == m){
    remove_rows_in_batches_(s, rows_to_remove)
    console.log(`removed ${rows_to_remove.length} rows from ${s.getName()}`)
  }else{
    m(s,rows_to_remove)
  }
}

function is_unwanted_(row){
    const raw_values = row.getValues()[0];
    const values = raw_values.filter(Boolean);
    if (values.length <= 1){
      return true
    }
    if (values.length > 3 && values[0] == "Earnings" && values[1] == "Begin Date"	&& values[2] == "End Date"){
      return true
    }
    if(values[0].startsWith("XXX") || values[0].startsWith("Direct Deposit")){
      return true
    }
    return false
}

function remove_rows_in_batches_(sheet, rows){
  const batches = gather_rows_(rows)
  for (const row_batch of batches){
    sheet.deleteRows(row_batch.start, row_batch.howMany)
  }
}

function gather_rows_(rows){
  const batches = []
  for (const current_row  of rows){
    let current_set = -1
    if (batches.length>0){
      current_set = batches[batches.length-1].start
    }
    if (current_set-1 == current_row){
      batches[batches.length-1].start = current_row
      batches[batches.length-1].howMany++
    } else {
      batches.push({start:current_row, howMany: 1})
    }
  }
  return batches
}


function remove_rows_1by1_(sheet, rows){
  for (const row_num of rows){
    sheet.deleteRow(row_num)
  }
}
