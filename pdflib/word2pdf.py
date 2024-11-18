from pathlib import Path
import win32com.client

def convert_word_to_pdf(in_file:str | Path, out_file:str | Path) -> None:
      '''Takes the in_file and converts it in the out_file. 
         Only takes absolute path
         The input strings must be raw and the path seperator must be : \
         Returns None. 
         '''
      wdFormatPDF = 17

      word = win32com.client.Dispatch('Word.Application') # Open the word application
      doc = word.Documents.Open(in_file) # Open the file to convert
      
      doc.SaveAs(out_file, FileFormat=wdFormatPDF) # Save the converted file
      # Close the docx file and the application
      doc.Close()
      word.Quit()

