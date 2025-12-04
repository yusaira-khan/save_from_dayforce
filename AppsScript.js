function main() {
  const sheets = SpreadsheetApp.getActiveSpreadsheet().getSheets();
  for (let s of sheets.slice(0)){
    remove_duplicates_(s)
  }
}

function remove_unwanted_(s, m=null){
  const range = s.getDataRange()
  const num_rows = range.getNumRows()
  const rows_to_remove = []

  for (let offset=0; offset<num_rows; offset++){
    const row = range.offset(offset, 0, 1);
    if (is_unwanted_(row)){
      rows_to_remove.push(row.getRow())
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
  const batches = gather_rows_(sanitize_rows_(rows))
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
  
  for (const row_num of sanitize_rows_(rows)){
    sheet.deleteRow(row_num)
  }
}
function sanitize_rows_(rows){
  const unique = [...new Set(rows)]
  unique.sort().reverse() // descending order
  return unique
}

function remove_duplicates_(s){
  const range = s.getDataRange()
  const num_rows = range.getNumRows()
  const range_values = get_values_(range)
  const rows_to_remove = []
  for (let offset=0; offset<num_rows; offset++){
    const rowNum = range.offset(offset, 0, 1).getRow();

    if (is_duplicate_in_(range_values, offset)){
      rows_to_remove.push(rowNum)
    }
  }
  remove_rows_in_batches_(s, rows_to_remove)
  console.log(`removed ${rows_to_remove} rows from ${s.getName()}`)
}

function is_duplicate_in_(all_row_values, test_row_idx){
  const test_row = all_row_values[test_row_idx]
  const duplicates = all_row_values.filter((candidate, candidate_idx)=> candidate_idx != test_row_idx && candidate[0]==test_row[0] )
  if (duplicates.length == 0){
    return false
  } else if(duplicates.length > 1) {
    throw new Error(`Unhandled! ${duplicates.length} possible duplicates found for ${test_row_idx}! under_test=${JSON.stringify(test_row)},canidates=${JSON.stringify(duplicates)} `)
  }else {
    const candidate = duplicates[0];
    if (test_row.length < candidate.length){
      return is_subrow_(candidate, test_row)
    }
    if (test_row.length == candidate.length){
      throw new Error(`Unhandled! ${test_row.length} has exact length match! under_test=${JSON.stringify(test_row)} canidate=${JSON.stringify(candidate)}`)
    }
    return false
  }
  //is_duplicate_in_([[1,2,3],[1]],1) => true
  //is_duplicate_in_([[1,2,3],[1]],0) => false
  //is_duplicate_in_([[1,2,3],[0,2,3]],0) => false
}

function is_subrow_(big, small){
  for(let idx=0; idx<small.length; idx++){
    if(big[idx]==small[idx]){
      continue;
    }else{
      return false
    }
  }
  return true

  // is_subrow_([1,2,3], [1]) => true
  // is_subrow_([1,2,3], [1,2,3]) => true
  // is_subrow_([0,1,2,3], [1]) => false
}

function get_values_(range){
  const r=range.getValues()
  const r2 = r.map((c)=>c.filter(Boolean))
  return r2
}
