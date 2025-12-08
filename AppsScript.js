function main() {

  const sheets = SpreadsheetApp.getActiveSpreadsheet().getSheets();
  const start_slice = 0
  let idx = 0
  for (let s of sheets.slice(start_slice)){
    console.log(`[${idx+start_slice}] reformatting ${s.getName()}`)

    reformat_table_(s)
    idx++;
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

function reformat_table_slow_(s){
  const range = s.getDataRange()
  const num_rows = range.getNumRows()
  const num_cols = range.getNumRows()

  for (let r=0; r<num_rows; r++){
    const row = range.offset(r, 0, 1);
    handle_aggregate_(row);
    for (let c=0; c<num_cols; c++){
      const cell = range.offset(r, c, 1,1)
      handle_number_format_(cell)
    }
    handle_bad_row_(row)
  }
}

const COLOR_LIGHT_CYAN_2 = "#a2c4c9"
const COLOR_GRAY = "#cccccc"

const AGGREGATE_COLOR = COLOR_LIGHT_CYAN_2
const AGGREGATE_NAMES = new Set(["Earnings", "Taxable Benefits","Taxes","Net Pay","Pre-Tax Deductions","Post-Tax Deductions","Reimbursements","Memo Information"])

function handle_aggregate_(range){
  if (AGGREGATE_NAMES.has(range.getValue())){
    range.setBackground(AGGREGATE_COLOR)
  }
}

const DOLLAR_FORMAT = '"$"#,##0.00;"$"\(#,##0.00\);$0.00;@'
const DECIMAL_4_REGEX = /^\d+\.\d{4}$/
const DECIMAL_4_FORMAT = "0.0000"
const DECIMAL_2_REGEX = /^\d+\.\d{2}$/
const DECIMAL_2_FORMAT = "0.00"

function handle_number_format_(cell){
  const value = cell.getValue()
  if (cell.isBlank() || typeof(value) === "number"){
    return;
  } else if (value.indexOf("$")!=-1){
    cell.setNumberFormat(DOLLAR_FORMAT)
  } else if (DECIMAL_4_REGEX.test(value)){
    cell.setNumberFormat(DECIMAL_4_FORMAT)
  } else if (DECIMAL_2_REGEX.test(value)){
    cell.setNumberFormat(DECIMAL_2_FORMAT)
  }
}

function handle_bad_row_(row){
  const values = row.getValues()[0]
  if (values.length>=5){
    if (values[0]=='' && values[1]=='' && values[3]=='Amount' && values[4]=='' && values[5]=='Amount' ){
      row.getSheet().deleteRow(row.getRow())
    }
  }
}

function reformat_table_(s){
  const table_range = get_proper_table_(s.getDataRange())
  adjust_column_formats_(table_range)

  const num_rows = table_range.getNumRows()
  for (let r=0; r<num_rows; r++){
    const row = table_range.offset(r, 0, 1);

    handle_aggregate_(row);
  }

  // format_main_header_(table_range.offset(0, 0, 1))
}

const DOLLAR_FORMAT = '"$"#,##0.00;"$"\(#,##0.00\);$0.00'
const DECIMAL_4_FORMAT = "0.0000"
const DECIMAL_2_FORMAT = "0.00"

function adjust_column_formats_(table_range){
  const s= table_range.getSheet()
  const num_rows = table_range.getNumRows()
  const num_columns= table_range.getNumColumns();
  s.setColumnWidth(1,150)
  s.setColumnWidths(2, num_columns-1, 100)
  table_range.offset(0, 1,num_rows, 1).setNumberFormat(DECIMAL_2_FORMAT)
  table_range.offset(0, 2, num_rows,1).setNumberFormat(DECIMAL_4_FORMAT)
  table_range.offset(0, 3, num_rows,1).setNumberFormat(DOLLAR_FORMAT)
  table_range.offset(0, 4, num_rows, 1).setNumberFormat(DECIMAL_2_FORMAT)
  table_range.offset(0, 5, num_rows,1).setNumberFormat(DOLLAR_FORMAT)
}

function format_main_header_(row){
  if (!row.getCell(1,2).isPartOfMerge()){
    if (!row.getCell(1,3).isBlank()){
      row.getCell(1,5).setValue(row.getCell(1,3).getValue())
    }
    row.offset(0,1,1,3).merge()
    row.offset(0,3,1,2).merge()
  }
}

function get_proper_table_(data_range){
  const num_columns = data_range.getNumColumns()
  const sheet = data_range.getSheet()

  const main_header_row = find_main_header_(data_range)
  const last_header_row = find_last_header_(data_range)

  const num_table_rows = 1 + last_header_row.getRow() - main_header_row.getRow()
  if (main_header_row.getRow() > 1){
    const table_range_spec = sheet.getRange(main_header_row.getRow(), 1, num_table_rows+3)
    sheet.moveRows(table_range_spec, 1)
  }
  const table_range = sheet.getRange(1,1,num_table_rows, num_columns)
  return table_range
}

const MAIN_HEADER_TEXT_OLD = "Detail at the time pay statement issued"
const MAIN_HEADER_TEXT_REPLACE = "Pay Statement"

function find_main_header_(range){
  for (let r=0; r<range.getNumRows(); r++){
    const row = range.offset(r, 0, 1);
      if (row.getValue()==MAIN_HEADER_TEXT_OLD ){
        row.getCell(1,1).setValue(MAIN_HEADER_TEXT_REPLACE)
        return row
      } else if (row.getValue()==MAIN_HEADER_TEXT_REPLACE){
        return row
      }
  }
  throw new Error("Unhandled!")
}

function find_last_header_(range){
  for (let r=0; r<range.getNumRows(); r++){
    const row = range.offset(r, 0, 1);
      if (row.getValue()=="Net Pay"){
        handle_bad_row_(row.offset(-1,0))
        return row
      }
  }
  throw new Error("Unhandled!")
}

