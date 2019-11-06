list.of.packages <- c("dplyr", "tidyr", "gridExtra", "openxlsx")
new.packages <- list.of.packages[!(list.of.packages %in% installed.packages()[,"Package"])]
if(length(new.packages)) install.packages(new.packages)
library(dplyr)
library(tidyr)          # for separate()
library(gridExtra)      # for grid.arrange & tableGrob
library(openxlsx)       # for Excel (but does not read xls files)

outfile <- readline("Specify output file (e.g. FINA1000_2017Fall_Final): ")
sections <- strsplit(readline("Specify section (e.g. ABC or CE): "), split="")[[1]]
courses <- paste0(readline("Specify Course Name (fixed part of Banner xls filenames, e.g. FINA_1000_1): "), sections)
subset <- readline("Specify set of exam numbers (e.g. Odds, Evens or leave blank for All): ")
subset <- ifelse(subset=="","All", subset)
seed <- readline("Enter numeric seed for reproducable sorting or leave empty: ")
CompileTeXFlag <- ifelse(readline("Enter 0 to NOT compile LaTeX files automatically, otherwise leave empty: ")==0,FALSE,TRUE)
A4Flag <- ifelse(readline("Enter a4 for A4 paperformat, or leave empty for letter format: ")==0,FALSE,TRUE)

#Default (test) values
outfile <- ifelse(outfile=="",'FINA1000_2017Fall_Final', outfile)
if (length(sections)==0) {
  sections <- c('A','B','C')
}
if (length(courses)==0) {
  courses <- paste0('FINA_1000_1', sections)
}

# create masterlist
courselist <- list()
for (i in 1:length(courses)) {
    courselist[[i]] <- read.csv(paste0(courses[i],'.csv'),strip.white = T)
    courselist[[i]]$Course <- courses[i]
}
masterlist <- do.call('rbind',courselist)
masterlist <- masterlist[complete.cases(masterlist),]
print(masterlist$Student.Name)
masterlist$Student.Name <- as.character(masterlist$Student.Name)
if (!is.na(as.numeric(seed))) {
  seed <- as.numeric(seed)
  set.seed(seed)
  print(paste0("This sorting can be reproduced with seed: ", seed))
} else if (seed=="") {
  print("No seed selected: Generating non-reproducible sorting!")
} else {
  print("Seed not numeric: Generating non-reproducible sorting!")
}
masterlist <- masterlist %>%
    separate(Student.Name, c("LastName","FirstName"), ',') %>%  # split Student Name
    mutate(foo = runif(nrow(masterlist)))  %>%                  # assign random number
    mutate(foo = ifelse(Record.Number == 999, 2, foo))  %>%     # mark Disability Center Students
    arrange(foo) %>%                                            # sort by random number
    mutate(ExamNumber = 1:nrow(masterlist))                     # assign ExamNumbers
    if (subset == "Odds") {
      masterlist <- masterlist %>% mutate(ExamNumber = 2*ExamNumber - 1) # use only odd ExamNumbers
    } else if (subset == "Evens") {
      masterlist <- masterlist %>% mutate(ExamNumber = 2*ExamNumber) # use only even ExamNumbers
    }

# Display warning if there are students with identical Last & First Name
if (dim(masterlist[duplicated(masterlist[c("LastName","FirstName")]),])[1] > 0) {
  warning("Masterlist contains students with identical First AND Last Name")
  masterlist[duplicated(masterlist[c("LastName","FirstName")]),]
}

# Function to write ExamNumber lookup lists to pdf files
#paper <- ifelse(A4Flag,"a4","letter") # if used tables appear too small
write2pdf <- function(df, rowsPerPage, outfile, filesuffix) {
  maxnrow = nrow(df)
  npages = ceiling(maxnrow/rowsPerPage)
  #pdf(file = paste0(outfile,"-",filesuffix,".pdf"), paper = paper, pointsize = 16)
  pdf(file = paste0(outfile,"-",filesuffix,".pdf"))
  for (i in 1:npages) {
    idx = seq(1+((i-1)*rowsPerPage), min(i*rowsPerPage,maxnrow))
    #grid.arrange(tableGrob(df[idx,], rows = NULL, theme = ttheme_default(base_size = 16) ), top = outfile)
    grid.arrange(tableGrob(df[idx,], rows = NULL), top = outfile)
  }
  dev.off()
}
# Create ExamNumber lookup lists
listByLast <- masterlist %>% select(LastName, FirstName, ExamNumber) %>% arrange(LastName, FirstName)
write2pdf(listByLast, 22, outfile, "ExamListByLastName")

listByA <- masterlist %>% select(ID, ExamNumber) %>% arrange(as.character(ID))
write2pdf(listByA, 22, outfile, "ExamListByA")

# Create Excel file
xlsxlist <- list("Masterlist" = masterlist, "by Name" = listByLast, "by A Number" = listByA)
write.xlsx(xlsxlist, paste0(outfile,"-ExamList.xlsx"), row.names=F)
print("Masterfile successfully created!")
print("Examlist (By Last Name) successfully created!")
print("Examlist (By A Number) successfully created!")

# Create Signature Sheet
m <- masterlist
sigfile <- paste0(outfile,"-signaturesheet.tex")
if (A4Flag) {
  geometry.package <- "\\usepackage[a4paper,margin=3mm]{geometry}"
  table.header <- "\\begin{tabular}{|c|L{8.5cm}|C{3cm}|C{4.3cm}|}"
} else {
  geometry.package <- "\\usepackage[margin=3mm]{geometry}"
  table.header <- "\\begin{tabular}{|c|L{8.5cm}|C{3cm}|C{5cm}|}"
}
write("\\documentclass[12pt]{article}", file=sigfile)
write(geometry.package, file=sigfile, append = T)
write("\\usepackage{array}", file=sigfile, append = T)
write("\\usepackage{fancyhdr}", file=sigfile, append = T)
write("\\usepackage{lastpage}", file=sigfile, append = T)
write("\\newcolumntype{L}[1]{>{\\raggedright\\let\\newline\\\\\\arraybackslash\\hspace{0pt}}m{#1}}",file=sigfile, append = T)
write("\\newcolumntype{C}[1]{>{\\centering\\let\\newline\\\\\\arraybackslash\\hspace{0pt}}m{#1}}",file=sigfile, append = T)
write("\\newcolumntype{R}[1]{>{\\raggedleft\\let\\newline\\\\\\arraybackslash\\hspace{0pt}}m{#1}}",file=sigfile, append = T)
write("\\pagestyle{fancy}", file=sigfile, append = T)
if (is.numeric(seed)) {
  write(paste0("\\lhead{",gsub("_","\\\\_",outfile)," (",seed,")}"), file=sigfile, append = T)
} else {
  write(paste0("\\lhead{",gsub("_","\\\\_",outfile),"}"), file=sigfile, append = T)
}
write("\\rhead{Page \\thepage\\ of \\pageref{LastPage}}", file=sigfile, append = T)
write("\\topmargin -20mm", file=sigfile, append = T)

write("\\begin{document}", file=sigfile, append = T)
write("\\begin{LARGE}", file=sigfile, append = T)

for (i in 1:nrow(m)) {
    if (i%%20 == 1) {
        if (i != 1) write("\\newpage", file=sigfile, append = T)
        write(table.header, file=sigfile, append = T)
        write("\\hline", file=sigfile, append = T)
        write("Exam & Name & A Number & Signature \\\\ \\hline", file=sigfile, append = T)
    }
    write(paste0(m$ExamNumber[i],"&",m$LastName[i],",",m$FirstName[i],"& & \\\\ \\hline"), file=sigfile, append = T)
    if (i%%20 == 0) write("\\end{tabular}", file=sigfile, append = T)
}
if (i%%20 != 0) write("\\end{tabular}", file=sigfile, append = T)
write("\\end{LARGE}", file=sigfile, append = T)
write("\\end{document}", file=sigfile, append = T)

if (CompileTeXFlag) {
  command <- paste0("pdflatex -synctex=1 -shell-escape -interaction=batchmode ",sigfile)
  if (Sys.info()['sysname']=="Linux") {
    command <- paste0("/usr/local/texlive/2019/bin/x86_64-linux/",command, "> /dev/null")
  }
  system(command)
  system(command)
  # Clean up LaTeX build files
  sigfileshort <- substr(sigfile, 1, nchar(sigfile)-4)
  system(command = paste0("rm -f ",sigfileshort, ".aux"))
  system(command = paste0("rm -f ",sigfileshort, ".log"))
  system(command = paste0("rm -f ",sigfileshort, ".synctex.gz"))
  print("Signature sheet successfully created!")
}

# Display warning if there are students with identical Last & First Name
if (dim(masterlist[duplicated(masterlist[c("LastName","FirstName")]),])[1] > 0) {
  warning("Masterlist contains students with identical First AND Last Name")
  masterlist[duplicated(masterlist[c("LastName","FirstName")]),]
}