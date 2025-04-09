######
# TITLE: UNECE Abstracts from Excel Records
######

### Setup start ###
# Install packages if not already installed
{
  list.of.packages <- c("tidyverse",
                        "readxl",
                        "janitor",
                        "officedown",
                        "tools"
                        )
  new.packages <- list.of.packages[!(list.of.packages %in% installed.packages()[,"Package"])]
  if(length(new.packages)) install.packages(new.packages)
  rm(list.of.packages,new.packages)
}

# Load packages
{
  library(tidyverse)
  library(readxl)
  library(janitor)
  library(officedown)
  library(tools)
}

### Setup end ###

# Begin by reading in the Excel data into a dataframe (df)
df <- read_excel(file.choose(),
                 sheet = 1) %>% # This assumes that the first sheet is the right one
# Clean the column names
clean_names() %>%
# Rename the columns in the df so that they match those in the "render_function" below.
rename(
  "country" = "country",
  "abstract" = "abstract_text",
  "abstract_title" = "abstract_title",
  "coauthors" = "co_authors",
  "organisation" = "organisation",
  "session_no" = "section"
)

# Clean up the first names
df$first_name = str_to_title(df$first_name)

# Clean up the surnames
df$last_name <- map_if(df$last_name,
       !grepl("[a-z]", df$last_name),
       str_to_title
)
       

# Meeting Organisational Data
# L-> Change these values to change them in the outputted MS Word docs.
{
  expertmeeting <- "Statistical Data Editing"
  venue <- "Vienna Austria"
  noticedate <- format(Sys.Date(), format="%d %B %Y")
  eventdate <- "7-9 October 2024"
}

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
                            coauthors,
                            session_no)
{
  rmarkdown::render("abst2word_SDE2024.Rmd", # This is the RMarkdown file
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
                                  coauthors = coauthors,
                                  session_no = session_no
                    ),
                    output_dir = "docs/SDE2024", # This is where the MS Word documents are output
                    # For now, I've added a sequential no to the output title to handle multiple submissions by the same authors
                    output_file=dynamname
  )  # Modify the code in output_file to change the file name formats for the MS Word docs
}

#This loop goes through the dataframe from your spreadsheet file and creates a PDF from the data in the relevant fields from each row.
x <- 2
for(i in 1:nrow(df)) {
  files_list <- list.files(
    path = "docs/SDE2024"
  ) %>%
    file_path_sans_ext()
  
  dynamname <- paste0("SDE2024_",df$session_no[i],"_", df$country[i], "_", df$last_name[i], "_A")
  
  # This checks for if a file by the same author already exists, and adds a sequential suffix number
  # to the file name if it does already exist
  if (dynamname %in% files_list) { 
    dynamname <- paste0(dynamname, "_No",x)
    x <- x + 1
  }
  
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
                  df$coauthors[i],
                  df$session_no[i]
  )
}
