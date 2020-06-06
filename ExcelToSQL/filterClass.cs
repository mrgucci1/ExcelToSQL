
using System.Linq;
using System.Data;


namespace ExcelToSQL
{
    class filterClass
    {
        public string[] filter(string[] files, string fileExten1, string fileExten2, string fileName)
        {
            //counter for new filtered array
            int counter = 0;
            //new array to store all filtered filed
            string[] filteredFiles = new string[files.Length];
            //loop through all files, only keeping .csv files
            for(int i = 0;i < files.Length;i++)
            {
                if(files[i].Contains(fileExten1) || files[i].Contains(fileExten2))
                {
                    if (files[i].Contains(fileName))
                    {
                        filteredFiles[counter] = files[i];
                        counter++;
                    }
                }
            }
            //remove null or whitespace from array
            filteredFiles = filteredFiles.Where(x => !string.IsNullOrWhiteSpace(x)).ToArray();
            return filteredFiles;
        }

        public string[] filterNoName(string[] files, string fileExten1, string fileExten2)
        {
            //counter for new filtered array
            int counter = 0;
            //new array to store all filtered filed
            string[] filteredFiles = new string[files.Length];
            //loop through all files, only keeping .csv files
            for (int i = 0; i < files.Length; i++)
            {
                if (files[i].Contains(fileExten1) || files[i].Contains(fileExten2))
                {
                    //if (files[i].Contains(fileName))
                    //{
                        filteredFiles[counter] = files[i];
                        counter++;
                    //}
                }
            }
            //remove null or whitespace from array
            filteredFiles = filteredFiles.Where(x => !string.IsNullOrWhiteSpace(x)).ToArray();
            return filteredFiles;
        }
    }
}
