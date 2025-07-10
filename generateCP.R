library(officer)
library(lubridate)

# Define Kinyarwanda months
months_rw <- c(
  "Mutarama", "Gashyantare", "Werurwe", "Mata", "Gicurasi", "Kamena",
  "Nyakanga", "Kanama", "Nzeli", "Ukwakira", "Ugushyingo", "Ukuboza"
)

# Date info
today_date <- today()
current_month_num <- month(today_date)
next_month_num <- month(today_date %m+% months(1))
current_month_rw <- months_rw[current_month_num]
next_month_rw <- months_rw[next_month_num]
current_year <- year(today_date)

# Load and modify cover page
wd <- read_docx("coverpage_reference.docx")
wd <- body_replace_all_text(wd, old_value = "months", new_value = toupper(current_month_rw))
wd <- body_replace_all_text(wd, old_value = "next", new_value = toupper(next_month_rw))
wd <- body_replace_all_text(wd, old_value = "years", new_value = as.character(current_year))

# Add a page break at the end of cover
wd <- body_add_break(wd)  # Adds a page break after cover content

wd<-body_add_docx(wd, src = "CPI_Report_Kinyarwanda.docx")
# Save generated cover page
print(wd, target = "generated_cover.docx")

