######
# TITLE: UNECE Abstracts from Excel Records
######

### Setup start ###
# Install packages if not already installed
{
  list.of.packages <- c("tidyverse",
                        "readxl",
                        "janitor",
                        "officedown"
                        )
  new.packages <- list.of.packages[!(list.of.packages %in% installed.packages()[,"Package"])]
  if(length(new.packages)) install.packages(new.packages)
  rm(list.of.packages,new.packages)
}
# Load packages
library(tidyverse)
library(readxl)
library(janitor)
library(officedown)


### Setup end ###

# Begin by reading in the Excel data into a dataframe (df).
# L-> If using another Excel file, then modify the file name string below. 
df <- read_excel("data/Abstract submission for Expert Meeting on Dissemination & Communication of Statistics 2023(1-41).xlsx",
                 sheet = "Abstract list") %>%
  clean_names() %>%
  rename(
    "country" = "country_represented_if_national_organisation_skip_if_international_organisation_or_academia",
    "session" = "for_which_session_you_are_submitting_the_contribution_for_the_description_of_the_sessions_see_information_notice_1",
    "abstract" = "abstract_of_the_contribution_approximately_200_400_words",
    "abstract_title" = "title_of_the_contribution",
    "coauthors" = "names_of_co_authors_if_any"
  )

# Meeting Organisational Data
# L-> Change these values to change them in the outputted MS Word docs.
expertmeeting <- "meetingPlaceholder"
venue <- "venuePlaceholder"
noticedate <- "DD MMMM 20XX"
eventdate <- "DD MMMM 20XX"

#This function is for handling the knitting (i.e. creation from a RMarkdown file template) of the PDFs.
render_function <- function(abstract_title, # These function arguments match with the parametres in the RMarkdown file
                            first_name,
                            last_name,
                            organisation,
                            country,
                            email_address,
                            abstract, 
                            expertmeeting, 
                            venue, 
                            noticedate, 
                            eventdate,
                            coauthors)
{
  rmarkdown::render("abst2word_V2.Rmd", # This is the RMarkdown file
                    params = list(abstract_title = abstract_title,
                                  first_name = first_name,
                                  last_name = last_name,
                                  organisation = organisation,
                                  country = country,
                                  email_address = email_address,
                                  abstract = abstract,
                                  expertmeeting = expertmeeting,
                                  venue = venue,
                                  noticedate = noticedate,
                                  eventdate = eventdate,
                                  coauthors = coauthors
                    ),
                    output_dir = "docs", # This is where the MS Word documents are output
                    output_file=paste0("Abstract", " - ", first_name, " ", last_name, " - ", format(Sys.Date(), format="%Y_%m_%d"))
  )  # Modify the code in output_file to change the file name formats for the MS Word docs
}

#This loop goes through the dataframe from your spreadsheet file and creates a PDF from the data in the relevant fields from each row.
for(i in 1:nrow(df)) {
  render_function(df$abstract_title[i],
                  df$first_name[i], 
                  df$last_name[i], 
                  df$organisation[i], 
                  df$country[i],
                  df$email_address[i],
                  df$abstract[i],
                  expertmeeting,
                  venue,
                  noticedate,
                  eventdate,
                  df$coauthors[i]
  )
}
