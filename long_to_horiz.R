require(tidyverse)

trnsps <- function(vert, meta) {
  
  custom_parser <- function(string) {
    keys <- stringr::str_extract_all(string, "(?<=\\[).*?(?=:)", simplify = T)
    vals <- stringr::str_extract_all(string, "(?<=:).*?(?=\\])", simplify = T)
    names(vals) <- keys
    return(vals)
  }
  
  meta$kv <- map(meta$DVG, custom_parser)
  
  #### Name patients ####
  patient_db <- count(vert, PatientCode) %>% select(-n)
  
  #### Remove missing ####
  vert <- filter(vert, !is.na(Value))
  
  # Check structure
  # table(vert$VisitName)
  # table(vert$FormName)
  # table(vert$Question) 
  # table(paste0(vert$VisitName, vert$FormName, vert$Question))
  
  # Generate unique column name
  # vert$instance_short <- if_else(vert$FormInstance == 1, "", as.character(vert$FormInstance), "")
  vert_unic <- unite(vert, colname, VisitName, FormName, Question, FormInstance, sep = "___")
  
  # Keep only handy variables
  vert_unic_simple <- select(vert_unic, PatientCode, colname, DataType, Value) %>% 
    mutate(., colname = str_remove(colname, "\\s*_{3}1$"))
  
  # Variable order
  question_db <- unite(meta, colname, VISIT, FORM, QUESTION, sep = "___", remove = F) %>% 
    .$colname %>% trimws()
  
  # Check structure
  # table(vert$DataType) 
  by_type_db <- split(as_tibble(vert_unic_simple), vert_unic_simple$DataType) %>% 
    map(~ select(., -DataType))
  
  #### Strings ####
  string_db <- spread(by_type_db$String, colname, Value, convert = T)
  
  #### Numerics ####
  long_db <- spread(by_type_db$Long, colname, Value, convert = T)
  float_db <- spread(by_type_db$Float, colname, Value, convert = T)
  
  #### Dates ####
  date_db <- mutate(by_type_db$Date, Value = as.Date(Value)) %>% 
    spread(colname, Value)
  
  # remove col names to change
  old_partialdate_question_db <- by_type_db$PartialDate$colname %>% unique() %>% unlist()
  question_db <- setdiff(question_db, old_partialdate_question_db)
  
  # change format
  by_type_db$PartialDate <- separate(by_type_db$PartialDate, Value, c("Year", "Month", "Day"), sep = "-") %>% 
    gather("ymd", "value", Year, Month, Day) %>% 
    unite(colname, colname, ymd, sep = "__")
  
  # store
  partialdate_db <- spread(by_type_db$PartialDate, colname, value, convert = T)
  
  # add new col names
  new_partialdate_question_db <- names(partialdate_db[-1])
  question_db <- append(question_db, new_partialdate_question_db)
  
  #### Factors ####
  bylab_dbs <- separate(by_type_db$List, colname, c(NA, "FORM", "QUESTION", NA), 
           remove = F, sep = "___") %>% 
    left_join(meta) %>% 
    {split(., paste0(.$FORM, .$QUESTION))} %>% 
    map(~ rowwise(.) %>% 
          mutate(not_enc_Value = Value,
                 Value = factor(Value, levels = names(kv), labels = kv)))
  
  list_dbs <- map(bylab_dbs, ~ spread(., colname, Value)) %>% 
    map(~ select(., -FORM, -QUESTION, -ALIAS, -LABEL, -DATATYPE, -DECIMALS, -kv, -not_enc_Value, -DVG))
  
  factor_db <- purrr::reduce(list_dbs, left_join)
  
  #### Boolean ####
  bool_db <- mutate(by_type_db$Boolean, Value = as.logical(Value)) %>% 
    spread(colname, Value)
  
  # Gather all types
  db_list <- list(patient_db, string_db, long_db, float_db, date_db, partialdate_db, factor_db, bool_db)
  
  db <- purrr::reduce(db_list, left_join) #%>% select(order(names(.))) %>% select(PatientCode, any_of(question_db), everything())
  
  return(db)
}





