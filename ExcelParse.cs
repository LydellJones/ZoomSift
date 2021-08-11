using System;
using System.Collections;
using System.IO;
//TODO: REPLACE BREAKS WITH CONTINUE
namespace Consoletestwork
{
    public class ExcelParse
    {
        private string line;//gather's the file line
        private bool container;//contains the boolean true or false for Sold
        private bool containertwo; //see above
        private bool containerthree; //see above
        private bool containerfour;
        public string author;//result of the split, contains the messenger
        private string purchaseparse; //contains the last end of the second split
        public int resultindex = 0;// index of the resulting numbers char list
        public string newS;
        public string[] printableresults;
        private string[] parsingline;//raw string of the file line
        public char[] numbers = { '0', '1', '2', '3', '4', '5', '6', '7', '8', '9' }; //chars of 0-9 to compare with split alphanumeric chars
        private string[] prephaseauthor;//first split
        private string[] midphaseauthor;//second split
        public string[] stringset; //individual characters  of purchaseparse to be compared
        ArrayList resultingnumbers = new ArrayList(); //results of the gathered numbers
        
       
        /*
         * The basic goal of this is to manuever through a
         * .txt file. Upon finding a line with "sold" in it
         * i want to document the name of the person who sent
         * the message and what item number they have purchased.
         * 
         * The formatting for the message is usually as follows:
         * 
         * [TIMESTAMP] From SENDER to RECIPIANT: [MESSAGE]
         * 
         * And every line contains this format. However...
         * 
         * Extracting the "sold" is a little more complicated
         * as we have instructed the senders to send "sold" in
         * the chat in this format.
         * 
         * SOLD [ITEM NUMBER]
         * 
         * Unfortunately, they may put it in the reverse order or explain 
         * specifically what they want surrounding this key format in a
         * wall of text.
         * 
         * My main challenge is to extract these numbers so that these
         * numbers can be exported into an excel sheet in the format.
         * 
         * [AUTHOR] [ITEM NUMBER]
         * 
         */ 

        public void SiftStart(string directory)
        {
            try
            {
                StreamReader sr = new StreamReader(directory);// Creation of the file reader
                try //try catch as a failsafe
                {
                    while (!sr.EndOfStream) //while the file has not ended
                    {
                        line = sr.ReadLine(); //assign file line
                        container = line.Contains("sold"); //assign a bool if the line contains sold
                        containertwo = line.Contains("Sold");//see above
                        containerthree = line.Contains("SOLD");//see above
                        containerfour = line.Contains("GHIA");

                        if ((container == true || containertwo == true || containerthree == true) && containerfour == false)//if a line has "sold"
                        {
                            parsingline = line.Split(" : "); //split the line into authr/recpt and message
                            prephaseauthor = parsingline[0].Split("From");//cages in the author's name using "from" and "to"
                            midphaseauthor = prephaseauthor[1].Split("to");//see above
                            author = midphaseauthor[0].Trim();//author of message is assigned
                            purchaseparse = parsingline[1];//assigned message half of the first split
                            stringset = purchaseparse.Split(" ");//message is broken up into words
                            Charsifter();
                            printableresults = (string[])resultingnumbers.ToArray(typeof(string));
                            resultingnumbers.Clear();
                            Console.WriteLine(author);
                            foreach(string i in printableresults)
                            {
                                Console.WriteLine(i);
                            }
                        }
                    }
                }
                catch (Exception e)//catches every exception from the try in order to prevent a memory leak
                {
                    Console.WriteLine("An error has occured");
                    Console.WriteLine(e.Message);
                }
                finally//closes the file no matter what
                {
                    sr.Close();
                }
            }
            catch(Exception e)
            {
                Console.WriteLine("Something is wrong with the file");
                Console.WriteLine(e.Message);
            }
        }
        /*(called from 81) compares a list of characters 0-9 with
         * the next character in the message to see if it is a number.
         * If not it is skipped over, if it is, it is recorded in another 
         * array to be recombined. (TODO: not done yet)
         */
        public void Charsifter()
        {
            try
            {
                foreach (string s in stringset)
                {
                    foreach (char num in numbers)
                    {
                        newS = s.ToLower();
                        if (newS.Contains(num))
                        {
                            if (newS.Contains('s'))
                            {
                                newS = newS.Replace("sold", "");
                                newS = newS.Trim();
                            }
                            if (newS.StartsWith('#'))
                            {
                                newS = newS.Replace("#", "");
                                newS = newS.Trim();
                            }
                            if (newS.Contains('-')){
                                break;
                            }
                            if (newS.StartsWith('0'))
                            {
                                while (newS.StartsWith('0'))
                                {
                                    newS = newS.Remove(0,1);
                                }
                            }
                            resultingnumbers.Add(newS.TrimStart().TrimEnd());
                            break;
                        }
                    }
                }
            }
            catch(Exception e)
            {
                Console.WriteLine("error in charsifter");
                Console.WriteLine(e.Message);
                return;
            }
        }
    }
}

