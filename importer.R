# set wd

do_magic <- function(add_empty_pts = F) {
  
  #Clear existing data and graphics
  # rm(list=ls())
  # graphics.off()
  
  #Load libraries
  library(Hmisc)
  library(tidyverse)
  library(lubridate)
  library(haven)
  library(openxlsx)
  
  #read Data
  cat("\n - Reading *.csv")
  BASELINE=read.csv("BASELINE.csv", header=TRUE, sep=";", na.strings = c(".", "NA"))
  CRITERIA=read.csv("CRITERIA.csv", header=TRUE, sep=";", na.strings = c(".", "NA"))
  MS_HISTORY=read.csv("MS_HISTORY.csv", header=TRUE, sep=";", na.strings = c(".", "NA"))
  LYMPHOCYT=read.csv("LYMPHOCYT.csv", header=TRUE, sep=";", na.strings = c(".", "NA"))
  COM_ALLERGY=read.csv("COM_ALLERGY.csv", header=TRUE, sep=";", na.strings = c(".", "NA"))
  MED_VAX=read.csv("MED_VAX.csv", header=TRUE, sep=";", na.strings = c(".", "NA"))
  PRIOR_COVID=read.csv("PRIOR_COVID.csv", header=TRUE, sep=";", na.strings = c(".", "NA"))
  V_DAY=read.csv("V_DAY.csv", header=TRUE, sep=";", na.strings = c(".", "NA"))
  BLOOD=read.csv("BLOOD.csv", header=TRUE, sep=";", na.strings = c(".", "NA"))
  TC_VAX=read.csv("TC_VAX.csv", header=TRUE, sep=";", na.strings = c(".", "NA"))
  FU_1M=read.csv("FU_1M.csv", header=TRUE, sep=";", na.strings = c(".", "NA"))
  TC_18M=read.csv("TC_18M.csv", header=TRUE, sep=";", na.strings = c(".", "NA"))
  MEDICATIONS=read.csv("MEDICATIONS.csv", header=TRUE, sep=";", na.strings = c(".", "NA"))
  COVID=read.csv("COVID.csv", header=TRUE, sep=";", na.strings = c(".", "NA"))
  ADV_EV=read.csv("ADV_EV.csv", header=TRUE, sep=";", na.strings = c(".", "NA"))
  
  
  #setting factors
  cat("\n - Setting cathegorical variables *.csv")
  BASELINE$SEX=factor(BASELINE$SEX,levels=c("1","2"))
  BASELINE$ETHNIA=factor(BASELINE$ETHNIA,levels=c("1","2","3","4"))
  BASELINE$EMPL=factor(BASELINE$EMPL,levels=c("1","2","3","4","5","6","7","8","9","10"))
  CRITERIA$IC1=factor(CRITERIA$IC1,levels=c("0","1"))
  CRITERIA$IC2=factor(CRITERIA$IC2,levels=c("0","1"))
  CRITERIA$IC3=factor(CRITERIA$IC3,levels=c("0","1"))
  CRITERIA$IC4=factor(CRITERIA$IC4,levels=c("0","1"))
  CRITERIA$IC5=factor(CRITERIA$IC5,levels=c("0","1"))
  CRITERIA$EC1=factor(CRITERIA$EC1,levels=c("0","1"))
  CRITERIA$EC2=factor(CRITERIA$EC2,levels=c("0","1"))
  MS_HISTORY$MSTYPE=factor(MS_HISTORY$MSTYPE,levels=c("1","2","3"))
  MS_HISTORY$RELAPSE_3M=factor(MS_HISTORY$RELAPSE_3M,levels=c("0","1"))
  MS_HISTORY$RELAPSE_ACT=factor(MS_HISTORY$RELAPSE_ACT,levels=c("0","1"))
  MS_HISTORY$DMD_LAST=factor(MS_HISTORY$DMD_LAST,levels=c("1","2","3","4","5","6","7","8","9","10","11","12","13","14","15","99"))
  MS_HISTORY$PREV_DMD=factor(MS_HISTORY$PREV_DMD,levels=c("0","1"))
  MS_HISTORY$PREV_DMD1=factor(MS_HISTORY$PREV_DMD1,levels=c("2","3","4","5","6","7","8","9","10","11","12","13","14","15","99"))
  MS_HISTORY$METH=factor(MS_HISTORY$METH,levels=c("0","1"))
  LYMPHOCYT$LYMPHO.factor=factor(LYMPHOCYT$LYMPHO,levels=c("1","2","3","4","5","6"))
  COM_ALLERGY$CEREB=factor(COM_ALLERGY$CEREB,levels=c("1","2","3"))
  COM_ALLERGY$KIDN=factor(COM_ALLERGY$KIDN,levels=c("1","2","3"))
  COM_ALLERGY$LIVER=factor(COM_ALLERGY$LIVER,levels=c("1","2","3"))
  COM_ALLERGY$HEART=factor(COM_ALLERGY$HEART,levels=c("1","2","3"))
  COM_ALLERGY$DIAB=factor(COM_ALLERGY$DIAB,levels=c("1","2","3"))
  COM_ALLERGY$HIV=factor(COM_ALLERGY$HIV,levels=c("1","2","3"))
  COM_ALLERGY$HBV=factor(COM_ALLERGY$HBV,levels=c("1","2","3"))
  COM_ALLERGY$HEMAT=factor(COM_ALLERGY$HEMAT,levels=c("1","2","3"))
  COM_ALLERGY$HYPERT=factor(COM_ALLERGY$HYPERT,levels=c("1","2","3"))
  COM_ALLERGY$DEPRESS=factor(COM_ALLERGY$DEPRESS,levels=c("1","2","3"))
  COM_ALLERGY$MTUM=factor(COM_ALLERGY$MTUM,levels=c("1","2","3"))
  COM_ALLERGY$OTHER_DISEASE=factor(COM_ALLERGY$OTHER_DISEASE,levels=c("1","2","3"))
  COM_ALLERGY$ALLERGY=factor(COM_ALLERGY$ALLERGY,levels=c("0","1"))
  COM_ALLERGY$PRE_SIGNS=factor(COM_ALLERGY$PRE_SIGNS,levels=c("0","1"))
  MED_VAX$MED=factor(MED_VAX$MED,levels=c("0","1"))
  MED_VAX$VACCINES=factor(MED_VAX$VACCINES,levels=c("0","1"))
  PRIOR_COVID$PRIOR_COVID=factor(PRIOR_COVID$PRIOR_COVID,levels=c("0","1"))
  PRIOR_COVID$SEV_COV=factor(PRIOR_COVID$SEV_COV,levels=c("0","1"))
  PRIOR_COVID$HOSP_COV=factor(PRIOR_COVID$HOSP_COV,levels=c("0","1"))
  PRIOR_COVID$ICU_COV=factor(PRIOR_COVID$ICU_COV,levels=c("0","1"))
  PRIOR_COVID$SERO_IGG=factor(PRIOR_COVID$SERO_IGG,levels=c("1","2"))
  PRIOR_COVID$SERO_IGM=factor(PRIOR_COVID$SERO_IGM,levels=c("1","2"))
  V_DAY$VAX_NAME=factor(V_DAY$VAX_NAME,levels=c("1","2","3","4","5","99"))
  V_DAY$SIDE_EFF01=factor(V_DAY$SIDE_EFF01,levels=c("0","1"))
  V_DAY$SIDE_EFF02=factor(V_DAY$SIDE_EFF02,levels=c("0","1"))
  V_DAY$SIDE_EFF03=factor(V_DAY$SIDE_EFF03,levels=c("0","1"))
  V_DAY$SIDE_EFF04=factor(V_DAY$SIDE_EFF04,levels=c("0","1"))
  V_DAY$SIDE_EFF05=factor(V_DAY$SIDE_EFF05,levels=c("0","1"))
  V_DAY$SIDE_EFF06=factor(V_DAY$SIDE_EFF06,levels=c("0","1"))
  V_DAY$SIDE_EFF07=factor(V_DAY$SIDE_EFF07,levels=c("0","1"))
  V_DAY$SIDE_EFF08=factor(V_DAY$SIDE_EFF08,levels=c("0","1"))
  V_DAY$SIDE_EFF09=factor(V_DAY$SIDE_EFF09,levels=c("0","1"))
  BLOOD$BLOOD1_ANC=factor(BLOOD$BLOOD1_ANC,levels=c("0","1"))
  BLOOD$BLOODINTERM_ANC=factor(BLOOD$BLOODINTERM_ANC,levels=c("0","1"))
  BLOOD$BLOOD2_ANC=factor(BLOOD$BLOOD2_ANC,levels=c("0","1"))
  TC_VAX$AE_TCVAX=factor(TC_VAX$AE_TCVAX,levels=c("0","1"))
  TC_VAX$COVID_TCVAX=factor(TC_VAX$COVID_TCVAX,levels=c("0","1"))
  TC_VAX$TRT_TCVAX=factor(TC_VAX$TRT_TCVAX,levels=c("0","1"))
  FU_1M$DMD_CHANGE=factor(FU_1M$DMD_CHANGE,levels=c("0","1"))
  FU_1M$DMD_INTER=factor(FU_1M$DMD_INTER,levels=c("0","1"))
  FU_1M$NEW_DMD=factor(FU_1M$NEW_DMD,levels=c("0","1"))
  FU_1M$NEW_DMD1=factor(FU_1M$NEW_DMD1,levels=c("2","3","4","5","6","7","8","9","10","11","12","13","14","15","99"))
  FU_1M$RELAPSE_FUP=factor(FU_1M$RELAPSE_FUP,levels=c("0","1"))
  FU_1M$RELAPSE_ACT=factor(FU_1M$RELAPSE_ACT,levels=c("0","1"))
  FU_1M$AE_FU1=factor(FU_1M$AE_FU1,levels=c("0","1"))
  FU_1M$COVID_FU1=factor(FU_1M$COVID_FU1,levels=c("0","1"))
  FU_1M$TRT_FU1=factor(FU_1M$TRT_FU1,levels=c("0","1"))
  TC_18M$AE_TC18=factor(TC_18M$AE_TC18,levels=c("0","1"))
  TC_18M$COVID_TC18=factor(TC_18M$COVID_TC18,levels=c("0","1"))
  TC_18M$TRT_TC18=factor(TC_18M$TRT_TC18,levels=c("0","1"))
  COVID$CONF_COVID=factor(COVID$CONF_COVID,levels=c("1","2"))
  COVID$COVID_SEVERE=factor(COVID$COVID_SEVERE,levels=c("0","1"))
  COVID$HOSP_COVID=factor(COVID$HOSP_COVID,levels=c("0","1"))
  COVID$ICU_COVID=factor(COVID$ICU_COVID,levels=c("0","1"))
  COVID$IGG_RES=factor(COVID$IGG_RES,levels=c("1","2"))
  COVID$IGM_RES=factor(COVID$IGM_RES,levels=c("1","2"))
  ADV_EV$AE_SEV=factor(ADV_EV$AE_SEV,levels=c("1","2","3"))
  ADV_EV$AE_REL=factor(ADV_EV$AE_REL,levels=c("1","2","3","4","5"))
  ADV_EV$AE_SER=factor(ADV_EV$AE_SER,levels=c("0","1"))
  
  #setting factor levels
  cat("\n - Converting labels *.csv")
  levels(BASELINE$SEX)=c("male","female")
  levels(BASELINE$ETHNIA)=c("caucasian","black or african-american","asian","other")
  levels(BASELINE$EMPL)=c("any or not employed","workman ","office clerk","physician","nurse","healthcare assistant","unfit for work ","student","retired","other")
  levels(CRITERIA$IC1)=c("no","yes")
  levels(CRITERIA$IC2)=c("no","yes")
  levels(CRITERIA$IC3)=c("no","yes")
  levels(CRITERIA$IC4)=c("no","yes")
  levels(CRITERIA$IC5)=c("no","yes")
  levels(CRITERIA$EC1)=c("no","yes")
  levels(CRITERIA$EC2)=c("no","yes")
  levels(MS_HISTORY$MSTYPE)=c("relapsing remitting MS (RRMS)","secondary progressive MS (SPMS)","primary progressive MS (PPMS)")
  levels(MS_HISTORY$RELAPSE_3M)=c("no","yes")
  levels(MS_HISTORY$RELAPSE_ACT)=c("no","yes")
  levels(MS_HISTORY$DMD_LAST)=c("never treated","alemtuzumab","azatioprina","daclizumab","dimethyl fumarate","glatiramer acetate","fingolimod","interferon","methotrexate","mitoxantrone","natalizumab","ocrelizumab","rituximab","teriflunomide","cladribine","other")
  levels(MS_HISTORY$PREV_DMD)=c("no","yes")
  levels(MS_HISTORY$PREV_DMD1)=c("alemtuzumab","azatioprina","daclizumab","dimethyl fumarate","glatiramer acetate","fingolimod","interferon","methotrexate","mitoxantrone","natalizumab","ocrelizumab","rituximab","teriflunomide","cladribine","other")
  levels(MS_HISTORY$METH)=c("no","yes")
  levels(LYMPHOCYT$LYMPHO.factor)=c("Grade 0 (normal: > 1000/mm3)","Grade 1 (800-999/mm3)","Grade 2 (500-799/mm3)","Grade 3 (200-499/mm3)","Grade 4 (<20/mm3)","not available")
  levels(COM_ALLERGY$CEREB)=c("no","yes, not treated","yes, treated")
  levels(COM_ALLERGY$KIDN)=c("no","yes, not treated","yes, treated")
  levels(COM_ALLERGY$LIVER)=c("no","yes, not treated","yes, treated")
  levels(COM_ALLERGY$HEART)=c("no","yes, not treated","yes, treated")
  levels(COM_ALLERGY$DIAB)=c("no","yes, not treated","yes, treated")
  levels(COM_ALLERGY$HIV)=c("no","yes, not treated","yes, treated")
  levels(COM_ALLERGY$HBV)=c("no","yes, not treated","yes, treated")
  levels(COM_ALLERGY$HEMAT)=c("no","yes, not treated","yes, treated")
  levels(COM_ALLERGY$HYPERT)=c("no","yes, not treated","yes, treated")
  levels(COM_ALLERGY$DEPRESS)=c("no","yes, not treated","yes, treated")
  levels(COM_ALLERGY$MTUM)=c("no","yes, not treated","yes, treated")
  levels(COM_ALLERGY$OTHER_DISEASE)=c("no","yes, not treated","yes, treated")
  levels(COM_ALLERGY$ALLERGY)=c("no","yes")
  levels(COM_ALLERGY$PRE_SIGNS)=c("no","yes")
  levels(MED_VAX$MED)=c("no","yes")
  levels(MED_VAX$VACCINES)=c("no","yes")
  levels(PRIOR_COVID$PRIOR_COVID)=c("no","yes")
  levels(PRIOR_COVID$SEV_COV)=c("no","yes")
  levels(PRIOR_COVID$HOSP_COV)=c("no","yes")
  levels(PRIOR_COVID$ICU_COV)=c("no","yes")
  levels(PRIOR_COVID$SERO_IGG)=c("positive","negative")
  levels(PRIOR_COVID$SERO_IGM)=c("positive","negative")
  levels(V_DAY$VAX_NAME)=c("AstraZeneca","Johnson & Johnson","Moderna","Pfizer/BioNTec","Sputnik V","Other")
  levels(V_DAY$SIDE_EFF01)=c("no","yes")
  levels(V_DAY$SIDE_EFF02)=c("no","yes")
  levels(V_DAY$SIDE_EFF03)=c("no","yes")
  levels(V_DAY$SIDE_EFF04)=c("no","yes")
  levels(V_DAY$SIDE_EFF05)=c("no","yes")
  levels(V_DAY$SIDE_EFF06)=c("no","yes")
  levels(V_DAY$SIDE_EFF07)=c("no","yes")
  levels(V_DAY$SIDE_EFF08)=c("no","yes")
  levels(V_DAY$SIDE_EFF09)=c("no","yes")
  levels(BLOOD$BLOOD1_ANC)=c("no","yes")
  levels(BLOOD$BLOODINTERM_ANC)=c("no","yes")
  levels(BLOOD$BLOOD2_ANC)=c("no","yes")
  levels(TC_VAX$AE_TCVAX)=c("no","yes")
  levels(TC_VAX$COVID_TCVAX)=c("no","yes")
  levels(TC_VAX$TRT_TCVAX)=c("no","yes")
  levels(FU_1M$DMD_CHANGE)=c("no","yes")
  levels(FU_1M$DMD_INTER)=c("no","yes")
  levels(FU_1M$NEW_DMD)=c("no","yes")
  levels(FU_1M$NEW_DMD1)=c("alemtuzumab","azatioprina","daclizumab","dimethyl fumarate","glatiramer acetate","fingolimod","interferon","methotrexate","mitoxantrone","natalizumab","ocrelizumab","rituximab","teriflunomide","cladribine","other")
  levels(FU_1M$RELAPSE_FUP)=c("no","yes")
  levels(FU_1M$RELAPSE_ACT)=c("no","yes")
  levels(FU_1M$AE_FU1)=c("no","yes")
  levels(FU_1M$COVID_FU1)=c("no","yes")
  levels(FU_1M$TRT_FU1)=c("no","yes")
  levels(TC_18M$AE_TC18)=c("no","yes")
  levels(TC_18M$COVID_TC18)=c("no","yes")
  levels(TC_18M$TRT_TC18)=c("no","yes")
  levels(COVID$CONF_COVID)=c("between two doses","after the last")
  levels(COVID$COVID_SEVERE)=c("no","yes")
  levels(COVID$HOSP_COVID)=c("no","yes")
  levels(COVID$ICU_COVID)=c("no","yes")
  levels(COVID$IGG_RES)=c("positive","negative")
  levels(COVID$IGM_RES)=c("positive","negative")
  levels(ADV_EV$AE_SEV)=c("mild","moderate","severe")
  levels(ADV_EV$AE_REL)=c("definitely related ","probably related","potentially related","unlikely to be related","not relate")
  levels(ADV_EV$AE_SER)=c("no","yes")
  
  #setting labels
  label(BASELINE$DIC)="Date of written informed consent"
  label(BASELINE$DOV_V0)="Date of visit"
  label(BASELINE$PROV)="Italian province "
  label(BASELINE$SEX)="Sex"
  label(BASELINE$AGE)="Age [years]"
  label(BASELINE$HEIGHT)="Height [cm]"
  label(BASELINE$WEIGHT)="Weight [kg]"
  label(BASELINE$ETHNIA)="Ethnicity "
  label(BASELINE$EMPL)="Employment"
  
  label(CRITERIA$IC1)="Provision of appropriate informed consent"
  label(CRITERIA$IC2)="Patients eligible to the Covid-19 vaccination and consent to receive it"
  label(CRITERIA$IC3)="Stated willingness to comply with all study procedures. Patient ability to comprehend the full nature and purpose of the study, ability to cooperate with the investigators and to comply with the requirements of the entire study"
  label(CRITERIA$IC4)="Male or female, aged ≥ 18 years"
  label(CRITERIA$IC5)="Confirmed MS diagnosis"
  label(CRITERIA$EC1)="Known allergic reactions to components of the vaccine"
  label(CRITERIA$EC2)="Any concomitant disease (cancer, autoimmune diseases) requiring treatment with B-cell–targeted therapies (e.g., rituximab or OCR), lymphocyte-trafficking blockers, alemtuzumab, anti-CD4, cladribine, cyclophosphamide, mitoxantrone, azathioprine, mycophenolate mofetil, cyclosporine, methotrexate, total body irradiation, or bone marrow transplantation"
  
  label(MS_HISTORY$MS_DATE)="Date of MS diagnosis "
  label(MS_HISTORY$MSTYPE)="MS type "
  label(MS_HISTORY$EDSS)="Last available EDSS "
  label(MS_HISTORY$EDSS_DATE)="Date of last EDSS  "
  label(MS_HISTORY$RELAPSE_3M)="Relapse in the last three months "
  label(MS_HISTORY$RELAPSE_DATE)="If 'yes', start date "
  label(MS_HISTORY$RELAPSE_ACT)="Relapse still active "
  label(MS_HISTORY$DMD_LAST)="Last DMD "
  label(MS_HISTORY$DMD_LAST_OTH)="If 'other', specify "
  label(MS_HISTORY$DMD_START)="Start date of last DMD"
  label(MS_HISTORY$DMD_STOP)="If stopped, date of last dose of DMD"
  label(MS_HISTORY$DMD_ONG)="Ongoing"
  label(MS_HISTORY$REASON_STOP01)="for COVID"
  label(MS_HISTORY$REASON_STOP02)="end of the therapeutic cycle"
  label(MS_HISTORY$REASON_STOP03)="lack of efficacy"
  label(MS_HISTORY$REASON_STOP04)="patient's decision"
  label(MS_HISTORY$REASON_STOP05)="pregnancy"
  label(MS_HISTORY$REASON_STOP06)="pregnancy planning"
  label(MS_HISTORY$REASON_STOP07)="adverse event or side effect"
  label(MS_HISTORY$REASON_STOP08)="other"
  label(MS_HISTORY$REASON_OTH)="Other reason"
  label(MS_HISTORY$PREV_DMD)="Patient previously treated (before the last DMD)?"
  label(MS_HISTORY$PREV_DMD1)="If 'yes' report the name of the last previous DMD "
  label(MS_HISTORY$PREV_DMD1_OTH)="If 'other', specify "
  label(MS_HISTORY$METH)="Any cycle of Methylprednisolone or other glucocorticoids in the last month?"
  label(MS_HISTORY$METH_START)="From"
  label(MS_HISTORY$METH_STOP)="To"
  label(MS_HISTORY$METH_ONG)="Ongoing"
  
  label(LYMPHOCYT$LYMPHO)="Recent lymphocytes counts (within 3 months prior vaccination)"
  
  label(COM_ALLERGY$CEREB)="Cerebrovascular disease "
  label(COM_ALLERGY$KIDN)="Chronic kidney disease "
  label(COM_ALLERGY$LIVER)="Chronic liver disease "
  label(COM_ALLERGY$HEART)="Coronary heart disease "
  label(COM_ALLERGY$DIAB)="Diabetes "
  label(COM_ALLERGY$HIV)="HIV "
  label(COM_ALLERGY$HBV)="HBV "
  label(COM_ALLERGY$HEMAT)="Hematological disease "
  label(COM_ALLERGY$HYPERT)="Hypertension "
  label(COM_ALLERGY$DEPRESS)="Major depressive disorder "
  label(COM_ALLERGY$MTUM)="Malignant tumor "
  label(COM_ALLERGY$OTHER_DISEASE)="Other disease "
  label(COM_ALLERGY$ALLERGY)="History of significant allergic reactions "
  label(COM_ALLERGY$ALLERGY_SP)="If 'yes', specify the product(s) and reaction "
  label(COM_ALLERGY$PRE_SIGNS)="Pre-vaccination signs or symptoms in this last week (i.e. Cold, fever…) "
  label(COM_ALLERGY$SIGNS_SP)="If 'yes', specify "
  
  label(MED_VAX$MED)="Previous (started 3 months before) or concomitant medication? "
  label(MED_VAX$VACCINES)="Other vaccines administered in the last year? "
  label(MED_VAX$VACC1)="1. Name of vaccine"
  label(MED_VAX$VAX1_DATE)="1. Date of last dose"
  label(MED_VAX$VACC2)="2. Name of vaccine"
  label(MED_VAX$VAX2_DATE)="2. Date of last dose"
  label(MED_VAX$VACC3)="3. Name of vaccine"
  label(MED_VAX$VAX3_DATE)="3. Date of last dose"
  label(MED_VAX$VACC4)="4. Name of vaccine"
  label(MED_VAX$VAX4_DATE)="4. Date of last dose"
  label(MED_VAX$VACC5)="5. Name of vaccine"
  label(MED_VAX$VAX5_DATE)="5. Date of last dose"
  
  label(PRIOR_COVID$PRIOR_COVID)="Prior confirmed SARS-CoV-2 infection? "
  label(PRIOR_COVID$SEV_COV)="If 'yes', was it SEVERE? "
  label(PRIOR_COVID$HOSP_COV)="Hospitalization for Covid? "
  label(PRIOR_COVID$HOSP_DAY)="For n days "
  label(PRIOR_COVID$ICU_COV)="ICU admission for Covid? "
  label(PRIOR_COVID$ICU_DAY)="For n days "
  label(PRIOR_COVID$PCRPOS_DATE)="Date of last PCR-POSITIVE test "
  label(PRIOR_COVID$PCRNEG_DATE)="Date of last PCR-NEGATIVE test "
  label(PRIOR_COVID$SERO_DATE)="Date of last serological test, if any "
  label(PRIOR_COVID$SERO_IGG)="IgG "
  label(PRIOR_COVID$SERO_IGM)="IgM "
  label(PRIOR_COVID$MUSC_19)="If available, please report MUSC-19 ID (XXX-YYY)"
  
  label(V_DAY$VAX_NAME)="Vaccine product"
  label(V_DAY$DOSE1_DATE)="Date of first vaccine dose "
  label(V_DAY$DOSE2_DATE)="Date of second vaccine dose, if required "
  label(V_DAY$SIDE_EFF01)="Injection site pain"
  label(V_DAY$SIDE_EFF02)="Swelling / redness"
  label(V_DAY$SIDE_EFF03)="Asthenia / fatigue"
  label(V_DAY$SIDE_EFF04)="Headache"
  label(V_DAY$SIDE_EFF05)="Chills"
  label(V_DAY$SIDE_EFF06)="Myalgia"
  label(V_DAY$SIDE_EFF07)="Fever (greater than 37.5 ° C)"
  label(V_DAY$SIDE_EFF08)="Nausea, vomiting"
  label(V_DAY$SIDE_EFF09)="Intestinal disorders"
  label(V_DAY$SIDE_EFF_OTH)="Other, specify"
  
  label(BLOOD$BLOOD1_DATE)="Date of first blood sample collection "
  label(BLOOD$NO_BLOOD1)="Not executed"
  label(BLOOD$ANTIB_FIRST)="SARS-CoV-2-S antibodies [U/mL]"
  label(BLOOD$ANTIB_FIRST_2N)="SARS-CoV-2-N antibodies [U/mL]"
  label(BLOOD$BLOOD1_ANC)="Blood sample for immunological ancillary study "
  label(BLOOD$BLOODINT_DATE)="Date of intermediate sample collection "
  label(BLOOD$NO_BLOODINTERM)="Not executed"
  label(BLOOD$BLOODINTERM_ANC)="Blood sample for immunological ancillary study "
  label(BLOOD$BLOOD2_DATE)="Date of second blood sample collection "
  label(BLOOD$NO_BLOOD2)="Not executed"
  label(BLOOD$ANTIB_SECOND)="SARS-CoV-2-S antibodies [U/mL]"
  label(BLOOD$ANTIB_SECOND_2N)="SARS-CoV-2-N antibodies [U/mL]"
  label(BLOOD$BLOOD2_ANC)="Blood sample for immunological ancillary study "
  
  label(TC_VAX$DOV_TCVAX)="Date of contact "
  label(TC_VAX$AE_TCVAX)="Does the patient experience any adverse event since the baseline visit? "
  label(TC_VAX$COVID_TCVAX)="Does the patient experienced SARS-CoV-2 infection since the baseline visit? "
  label(TC_VAX$TRT_TCVAX)="Does the patient start a new treatment or modify an ongoing treatment since the baseline visit?"
  
  label(FU_1M$DOV_FU1)="Date of visit "
  label(FU_1M$DMD_CHANGE)="Any change in DMD compared to baseline visit "
  label(FU_1M$DMD_INTER)="Previous DMD interrupted "
  label(FU_1M$DATE_INTERR)="Date of interruption "
  label(FU_1M$REASON_INTERR01)="for COVID"
  label(FU_1M$REASON_INTERR02)="end of the therapeutic cycle"
  label(FU_1M$REASON_INTERR03)="lack of efficacy"
  label(FU_1M$REASON_INTERR04)="patient's decision"
  label(FU_1M$REASON_INTERR05)="pregnancy"
  label(FU_1M$REASON_INTERR06)="pregnancy planning"
  label(FU_1M$REASON_INTERR07)="adverse event or side effect"
  label(FU_1M$REASON_INTERR08)="other"
  label(FU_1M$OTHER_REASON)="If 'other', specify "
  label(FU_1M$NEW_DMD)="New DMD"
  label(FU_1M$DMD_STARTFU1)="Start date of new DMD "
  label(FU_1M$NEW_DMD1)="Name of new DMD "
  label(FU_1M$NEW_DMD1_OTH)="If 'other', specify "
  label(FU_1M$FROM_METH)="From"
  label(FU_1M$TO_METH)="To"
  label(FU_1M$RELAPSE_FUP)="Any new relapse "
  label(FU_1M$RELAPSE_DATE)="If 'yes', start date "
  label(FU_1M$RELAPSE_ACT)="Relapse still active "
  label(FU_1M$AE_FU1)="Does the patient experience any adverse event since the last phone contact? "
  label(FU_1M$COVID_FU1)="Does the patient experienced SARS-CoV-2 infection since the last phone contact? "
  label(FU_1M$TRT_FU1)="Does the patient start a new treatment or modify an ongoing treatment since the last phone contact?"
  
  label(TC_18M$DOV_TC18)="Date of visit "
  label(TC_18M$AE_TC18)="Does the patient experience any adverse event since the last visit? "
  label(TC_18M$COVID_TC18)="Does the patient experienced SARS-CoV-2 infection since the last visit? "
  label(TC_18M$TRT_TC18)="Does the patient start a new treatment or modify an ongoing treatment since the last visit?"
  
  label(MEDICATIONS$MED_NAME)="Name of medication "
  label(MEDICATIONS$MED_IND)="Indication "
  label(MEDICATIONS$MED_START)="From"
  label(MEDICATIONS$MED_STOP)="To"
  label(MEDICATIONS$MED_ONG)="Ongoing"
  
  label(COVID$CONF_COVID)="Confirmed Covid: "
  label(COVID$COVID_SEVERE)="Was it SEVERE? "
  label(COVID$HOSP_COVID)="Hospitalization for Covid? "
  label(COVID$HOSP_DAYS)="For days "
  label(COVID$ICU_COVID)="ICU admission for Covid? "
  label(COVID$ICU_DAYS)="For days "
  label(COVID$POS_DATE)="Date of first PCR-POSITIVE test "
  label(COVID$NEG_DATE)="Date of PCR-NEGATIVE test "
  label(COVID$SER_DATE)="Date of serological test, if any "
  label(COVID$IGG_RES)="IgG "
  label(COVID$IGM_RES)="IgM "
  
  label(ADV_EV$AE_DESCR)="AE description "
  label(ADV_EV$AE_SEV)="Severity"
  label(ADV_EV$AE_REL)="Relationship"
  label(ADV_EV$AE_START)="Start date"
  label(ADV_EV$AE_STOP)="Stop date"
  label(ADV_EV$AE_ONG)="Ongoing"
  label(ADV_EV$AE_SER)="Serious Adverse Event (SAE)"
  
  # check vert
  # count(CRITERIA, PatientCode, name = "c") %>% count(c)
  # count(BASELINE, PatientCode, name = "c") %>% count(c)
  # count(MS_HISTORY, PatientCode, name = "c") %>% count(c)
  # count(COM_ALLERGY, PatientCode, name = "c") %>% count(c)
  # count(MED_VAX, PatientCode, name = "c") %>% count(c)
  # count(PRIOR_COVID, PatientCode, name = "c") %>% count(c)
  # count(V_DAY, PatientCode, name = "c") %>% count(c)
  # count(BLOOD, PatientCode, name = "c") %>% count(c)
  # count(TC_VAX, PatientCode, name = "c") %>% count(c)
  # count(FU_1M, PatientCode, name = "c") %>% count(c)
  # count(TC_18M, PatientCode, name = "c") %>% count(c)
  # count(COVID, PatientCode, name = "c") %>% count(c)
  # 
  # count(MEDICATIONS, PatientCode, name = "c") %>% count(c)
  # count(ADV_EV, PatientCode, name = "c") %>% count(c)
  
  # add empty pts
  PATIENTS <- readr::read_delim("_PATIENTS.csv", delim = ";", escape_double = F, na = c(".", "NA"), trim_ws = T) %>% 
    suppressMessages() %>% 
    select(PatientID, PatientCode = Patient_code, EnrollDate, SiteID, Created, CreatedByID, LastUpdate, LastUpdateByID)
  
  # horizontalize
  FU_1M <- dplyr::rename(FU_1M, 
                         RELAPSE_DATE_FU_1M = RELAPSE_DATE,
                         RELAPSE_ACT_FU_1M = RELAPSE_ACT)
  
  silent_full_join <- function(x, y) {suppressMessages(full_join(x, y))} 
  
  cat("\n - Merging")
  horiz <- list(BASELINE, MS_HISTORY, LYMPHOCYT, COM_ALLERGY, MED_VAX, PRIOR_COVID, V_DAY, BLOOD, TC_VAX, FU_1M, CRITERIA) %>% 
    purrr::map(~ select(., -c("PatientStatus", "VisitID", "VisitCode", 
                              "VisitStatus", "VisitInstance", "FormID", "FormCode", "FormStatus", 
                              "FormInstance", "LastUpdate"))) %>% 
    purrr::reduce(silent_full_join)
  
  if(add_empty_pts) { horiz <- full_join(PATIENTS, horiz) }
  
  horiz <- mutate_if(horiz, is.numeric, ~ haven::labelled(., label = attr(., "label")))
  
  
  cat("\n - Converting dates")
  date_cols <- names(select(horiz, DIC, contains("DOV_"), contains("DATE"), matches("_START$"), matches("_STOP$")))
  
  impute_day <- function(d, m, y) {
    if (is.na(d)) { return(c("15", m, y)) } else { return(c(d, m, y))}
  }
  
  date_parser <- function(string_date) {
    if (is.na(string_date)) { return(NA_Date_) }
    
    if(!grepl("-|/", string_date)) {cat("problem parsing ", string_date); return(NA_Date_)}
    
    if(grepl("/", string_date)) { return(lubridate::dmy(string_date)) }
    if(grepl("-", string_date)) {
      elem <- str_split(string_date, "-", n = 3)[[1]]
      if (length(elem) == 3) {names(elem) <- c("y", "m", "d")}
      if (length(elem) == 2) {
        names(elem) <- c("y", "m")
        elem['d'] <- NA_Date_
      }
      if (length(elem) == 1) {
        names(elem) <- c("y")
        elem['d'] <- NA_Date_
        elem['m'] <- NA_Date_
      }
    }
    
    parsed_date <- as.data.frame(t(elem))
    # parsed_date
    filled_date <- impute_day(parsed_date$d, parsed_date$m, parsed_date$y)
    
    lubridate::dmy(sprintf("%s-%s-%s", filled_date[1], filled_date[2], filled_date[3]))
  }
  
  parse_date <- function(dates) {map(dates, date_parser) %>% do.call('c', .)}
  
  
  horiz <- horiz %>% 
    mutate_at(.vars = vars(all_of(date_cols)), 
              .funs = list(`_dateFormat` = parse_date))
  
  # double check
  # horiz %>% select(all_of(date_cols), all_of(paste0(date_cols, "__dateFormat"))) %>% select({order(names(.))}) %>% View()
  
  
  cat("\n  DONE!")
  return(horiz)
}




# export
# haven::write_sav(horiz, format(Sys.time(), "CovaxiMS_%d%b%Y_%He%M.sav"))
# openxlsx::write.xlsx(horiz, format(Sys.time(), "CovaxiMS_%d%b%Y_%He%M.xlsx"))


