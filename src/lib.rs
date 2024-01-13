use serde;
use std::{error::Error, io, collections::HashMap, fs, path};
use rust_xlsxwriter::*;

#[derive(Default, Debug, serde::Serialize, serde::Deserialize)]
struct AutodeskParameters{
    params: HashMap<String, Value>
}
impl AutodeskParameters{
    // Simple add parameter capability, will support something other than floats eventually. Functions will be next.
    fn add_param(&mut self, key: String, value: Value)-> Result<(), Box<dyn Error>>{
        // println!("Parameter: {}, {:?}", key, value);
        self.params.insert(key, value);
        Ok(())
    }
    // Initial Functionality here. 
    fn write_parameter_file(&self, parameter_file_name: &str)-> Result<(), Box<dyn Error>> {
        let string_format = Format::new();
        let decimal_format = Format::new().set_num_format("0.000");
        let mut workbook = Workbook::new(parameter_file_name);
        let worksheet = workbook.add_worksheet();
        let mut row = 0;
        worksheet.write_string(row, 0, "Parameter", &string_format).expect("Couldn't write.");
        worksheet.write_string(row, 1, "Value", &string_format).expect("Couldn't write.");
        worksheet.write_string(row, 2, "Units", &string_format).expect("Couldn't write.");
        worksheet.write_string(row, 3, "Comments", &string_format).expect("Couldn't write.");

        row = 1;
        for param in &self.params{
            worksheet.write_string(row, 0, &param.1.name, &string_format).expect("Couldn't write.");
            worksheet.write_number(row, 1, param.1.value, &decimal_format).expect("Couldn't write.");
            worksheet.write_string(row, 2, &param.1.unit, &string_format).expect("Couldn't write.");
            worksheet.write_string(row, 3, &param.1.comment, &string_format).expect("Couldn't write.");
            row+=1;
        }
        workbook.close().expect("Couldn't close.");
        
        Ok(())
    }
    // Kind of the opposite of my use case, but will support one day.
    fn read_parameter_file(&self, parameter_file_name: &str)-> Result<(), Box<dyn Error>> {      
        Ok(())
    }
}
// Simple f64 Value holder.
#[derive(Default, Debug, serde::Serialize, serde::Deserialize)]
struct Value{
    name: String,
    value: f64,
    unit: String,
    comment: String
}


// Tests
#[cfg(test)]
mod tests {
    use super::*;
    #[test]
    fn test(){
        let test_file_name = "autodesk_params_test.xlsx";
        let mut autodesk_parameters = AutodeskParameters::default();
        autodesk_parameters.add_param("length".to_string(), Value{name: "length".to_string(), value: 0.01, unit: "ul".to_string(), comment: "N/A".to_string()}).expect("Couldn't insert.");
        
        if let Err(err) = autodesk_parameters.write_parameter_file(test_file_name){
            println!("Error writing parameter file: {}", err);
            assert!(false);
        }
        else{
            println!("Succesfully wrote parameter file.");
            assert!(true);            
        }
    }
}