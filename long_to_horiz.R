require(tidyverse)

trnsps <- function(vert, meta) {
  vert %>%
    select(PatientCode, Question, FormName, FormInstance, VisitName, Value) %>%
    tibble() %>%
    unite("Q", -Value, -PatientCode, sep = "___") %>%
    {cnames <<- unique(.$Q); .} %>%
    spread(Q, Value, convert = T) %>%
    select(PatientCode, all_of(cnames))
}

# trnsps <- function(vert, meta) {
#   
#   custom_parser <- function(string) {
#     keys <- stringr::str_extract_all(string, "(?<=\\[).*?(?=:)", simplify = T)
#     vals <- stringr::str_extract_all(string, "(?<=:).*?(?=\\])", simplify = T)
#     names(vals) <- keys
#     return(vals)
#   }
#   
#   meta$kv <- map(meta$DVG, custom_parser)
#   
#   #### Name patients ####
#   patient_db <- count(vert, PatientCode) %>% select(-n)
#   
#   #### Remove missing ####
#   vert <- filter(vert, !is.na(Value))
#   vert <- mutate(vert, Question = trimws(Question))
#   
#   # Check structure
#   # table(vert$VisitName)
#   # table(vert$FormName)
#   # table(vert$Question) 
#   # table(paste0(vert$VisitName, vert$FormName, vert$Question))
#   
#   # Generate unique column name
#   # vert$instance_short <- if_else(vert$FormInstance == 1, "", as.character(vert$FormInstance), "")
#   vert_unic <- unite(vert, colname, Question, FormName, FormInstance, VisitName, sep = "___")
#   
#   # Keep only handy variables
#   vert_unic_simple <- select(vert_unic, PatientCode, colname, DataType, Value) %>% 
#     mutate(., colname = str_remove(colname, "\\s*_{3}1$"))
#   
#   # Variable order
#   # question_db <- unite(meta, colname, VISIT, FORM, QUESTION, sep = "___", remove = F) %>% 
#   #   .$colname %>% trimws() %>% map(~paste0(., c("", paste0("___", 1:10)))) %>% flatten_chr()
#   question_db <- unique(vert_unic_simple$colname)
#   
#   # Check structure
#   # table(vert$DataType) 
#   by_type_db <- split(as_tibble(vert_unic_simple), vert_unic_simple$DataType) %>% 
#     map(~ select(., -DataType))
#   
#   #### Strings ####
#   string_db <- spread(by_type_db$String, colname, Value, convert = T)
#   
#   #### Numerics ####
#   long_db <- spread(by_type_db$Long, colname, Value, convert = T)
#   float_db <- spread(by_type_db$Float, colname, Value, convert = T)
#   
#   #### Partial dates ####
#   # # change format
#   # by_type_db$PartialDate <- separate(by_type_db$PartialDate, Value, c("Year", "Month", "Day"), sep = "-") %>% 
#   #   gather("ymd", "value", Year, Month, Day) %>% 
#   #   unite(colname, colname, ymd, sep = "__")
#   # 
#   # # remove col names to change
#   # old_partialdate_question_db <- by_type_db$PartialDate$colname %>% unique() %>% unlist()
#   # which(question_db %in% old_partialdate_question_db)
#   # 
#   # # store
#   # partialdate_db <- spread(by_type_db$PartialDate, colname, value, convert = T)
#   # 
#   # # add new col names
#   # new_partialdate_question_db <- names(partialdate_db[-1])
#   # question_db <- append(question_db, new_partialdate_question_db)
#   
#   partialdate_db <- separate(by_type_db$PartialDate, Value, c("Year", "Month", "Day"), sep = "-") %>% 
#     filter(!is.na(Year)) %>% 
#     mutate(Day = ifelse(is.na(Day), 15, Day),
#            Month = ifelse(is.na(Month), 6, Month)) %>% 
#     mutate(Value = as.Date(sprintf("%s-%s-%s", Year, Month, Day))) %>% 
#     select(-Year, -Month, -Day) %>% 
#     spread(colname, Value)
#   
#   #### Dates ####
#   date_db <- mutate(by_type_db$Date, Value = as.Date(Value)) %>% 
#     spread(colname, Value)
#   
#   #### Factors ####
#   bylab_dbs <- separate(by_type_db$List, colname, c("QUESTION", "FORM", NA, NA), 
#            remove = F, sep = "___") %>% 
#     left_join(meta) %>% 
#     {split(., paste0(.$FORM, .$QUESTION))} %>% 
#     map(unique)
#   
#   encoded_bylab_dbs <- map(bylab_dbs, ~ rowwise(.) %>% 
#           mutate(not_enc_Value = Value,
#                  Value = factor(Value, levels = names(kv), labels = kv))) %>% 
#     map(unique) %>% 
#     group_by(PatientCode) %>% mutate()
#   
#   list_dbs <- map(encoded_bylab_dbs, ~ spread(., colname, Value)) %>% 
#     map(~ select(., -VISIT, -FORM, -QUESTION, -ALIAS, -LABEL, -DATATYPE, -DECIMALS, -kv, -not_enc_Value, -DVG)) %>% 
#     map(unique)
#   
#   factor_db <- purrr::reduce(list_dbs, full_join)
#   
#   #### Boolean ####
#   bool_db <- mutate(by_type_db$Boolean, Value = as.logical(Value)) %>% 
#     spread(colname, Value)
#   
#   # Gather all types
#   db_list <- list(patient_db, string_db, long_db, float_db, date_db, partialdate_db, factor_db, bool_db)
#   
#   db <- purrr::reduce(db_list, left_join) %>% 
#     select(sort(names(.))) %>% 
#     select(PatientCode, any_of(question_db), everything())
#   
#   return(db)
# }





