list.of.packages <- c("dplyr", "tidyr", "gridExtra", "openxlsx")
new.packages <- list.of.packages[!(list.of.packages %in% installed.packages()[,"Package"])]
if(length(new.packages)) install.packages(new.packages)
library(dplyr)
library(tidyr)          # for separate()
library(gridExtra)      # for grid.arrange & tableGrob
library(openxlsx)       # for Excel (but does not read xls files)

outfile <- readline("Specify output file (e.g. FINA1000_2017Fall_Midterm2): ")  
sections <- strsplit(readline("Specify section (e.g. ABC or CE): "), split="")[[1]]
courses <- paste0(readline("Specify Course Name (fixed part of Banner filenames, e.g. FINA_1000_1): "), sections)
subset <- readline("Specify set of exam numbers (e.g. Odds, Evens or leave blank for All): ")
subset <- ifelse(subset=="","All", subset)
CompileTeXFlag <- ifelse(readline("Enter 0 to NOT compile LaTeX files automatically, otherwise leave empty: ")==0,FALSE,TRUE)
#Default (test) values
outfile <- ifelse(outfile=="",'FINA1000_2017Fall_Final', outfile)
sections <- ifelse(length(sections)==0, c('A','B','C'), sections)
courses <- ifelse(courses=="", paste0('FINA_1000_1', sections), courses)
  
# create masterlist
courselist <- list()
for (i in 1:length(courses)) {
    courselist[[i]] <- read.csv(paste0(courses[i],'.csv'))
    courselist[[i]]$Course <- courses[i]
}
masterlist <- do.call('rbind',courselist)
masterlist <- masterlist[complete.cases(masterlist),]
masterlist$Student.Name <- as.character(masterlist$Student.Name)
masterlist <- masterlist %>%
    separate(Student.Name, c("LastName","FirstName"), ',') %>%  # split Student Name
    mutate(foo = runif(nrow(masterlist)))  %>%                  # assign random number
    mutate(foo = ifelse(Record.Number == 999, 2, foo))  %>%     # mark Disability Center Students
    arrange(foo) %>%                                            # sort by random number
    mutate(ExamNumber = 1:nrow(masterlist))                     # assign ExamNumbers
    if (subset == 'Odds') {
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
write2pdf <- function(df, rowsPerPage, outfile, filesuffix) {
  maxnrow = nrow(df)
  npages = ceiling(maxnrow/rowsPerPage)
  pdf(file = paste0(outfile,"-",filesuffix,".pdf"))
  for (i in 1:npages) {
    idx = seq(1+((i-1)*rowsPerPage), min(i*rowsPerPage,maxnrow))
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

# Create Signature Sheet
m <- masterlist
sigfile <- paste0(outfile,"-signaturesheet.tex")
write("\\documentclass[12pt]{article}", file=sigfile)
write("\\usepackage[margin=0mm]{geometry}", file=sigfile, append = T)
write("\\usepackage{array}", file=sigfile, append = T)
write("\\usepackage{fancyhdr}", file=sigfile, append = T)
write("\\usepackage{lastpage}", file=sigfile, append = T)
write("\\newcolumntype{L}[1]{>{\\raggedright\\let\\newline\\\\\\arraybackslash\\hspace{0pt}}m{#1}}",file=sigfile, append = T)
write("\\newcolumntype{C}[1]{>{\\centering\\let\\newline\\\\\\arraybackslash\\hspace{0pt}}m{#1}}",file=sigfile, append = T)
write("\\newcolumntype{R}[1]{>{\\raggedleft\\let\\newline\\\\\\arraybackslash\\hspace{0pt}}m{#1}}",file=sigfile, append = T)
write("\\pagestyle{fancy}", file=sigfile, append = T)
#write(paste0("\\lhead{",gsub("_","\\\\_",inputfileshort),"}"), file=sigfile, append = T)
write(paste0("\\lhead{",gsub("_","\\\\_",outfile),"}"), file=sigfile, append = T)
write("\\rhead{Page \\thepage\\ of \\pageref{LastPage}}", file=sigfile, append = T)
write("\\topmargin -20mm", file=sigfile, append = T)

write("\\begin{document}", file=sigfile, append = T)
write("\\begin{LARGE}", file=sigfile, append = T)

for (i in 1:nrow(m)) {
    if (i%%20 == 1) {
        if (i != 1) write("\\newpage", file=sigfile, append = T)
        write("\\begin{tabular}{|c|L{9cm}|C{3cm}|C{5cm}|}", file=sigfile, append = T)
        write("\\hline", file=sigfile, append = T)
        write("Exam & Name & A Number & Signature \\\\ \\hline", file=sigfile, append = T)
    }
#    print(m$ExamNumber[i])
#    print(m$LastName[i])
    write(paste0(m$ExamNumber[i],"&",m$LastName[i],",",m$FirstName[i],"& & \\\\ \\hline"), file=sigfile, append = T)
    if (i%%20 == 0) write("\\end{tabular}", file=sigfile, append = T)
}
if (i%%20 != 0) write("\\end{tabular}", file=sigfile, append = T)
write("\\end{LARGE}", file=sigfile, append = T)
write("\\end{document}", file=sigfile, append = T)

if (CompileTeXFlag) {
  system(command = paste0("pdflatex -synctex=1 -shell-escape -interaction=nonstopmode ",sigfile))
  system(command = paste0("pdflatex -synctex=1 -shell-escape -interaction=nonstopmode ",sigfile))

  # Cleaning up LaTeX build files
  sigfileshort <- substr(sigfile, 1, nchar(sigfile)-4)
  system(command = paste0("rm -f ",sigfileshort, ".aux"))
  system(command = paste0("rm -f ",sigfileshort, ".log"))
  system(command = paste0("rm -f ",sigfileshort, ".synctex.gz"))
}

# Display warning if there are students with identical Last & First Name
if (dim(masterlist[duplicated(masterlist[c("LastName","FirstName")]),])[1] > 0) {
  warning("Masterlist contains students with identical First AND Last Name")
  masterlist[duplicated(masterlist[c("LastName","FirstName")]),]
}

# opening file in SumatraPDF does not work yet
#system(command = paste0("C:/D-other/portable/SumatraPDF.exe ", sigfileshort, ".pdf  -reuse-instance", sigfileshort, ".pdf"))
#system(command = paste0("C:/D-other/portable/SumatraPDF.exe ", sigfileshort, ".pdf"))
#system(command = paste0("C:/D-other/portable/SumatraPDF.exe ", sigfileshort,".pdf"))

