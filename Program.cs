using System;
using System.Collections.Generic;
using System.Text;
using System.Windows.Automation;
using System.Windows;
using System.Threading;
using System.Windows.Forms;

namespace HelperProgramTest
{
    class Program
    {
        static void Main(string[] args)
        {
            //Directory of books
            string filelocation = "directory where the books can be found";

            //List of to be retyped books
            List<string> boeken = new List<string>();
            boeken.Add("title of book");
            boeken.Add("title of book");

            
            Console.WriteLine("Start test");

            foreach (var boek in boeken)
            {

                //Find and open notepad

                AutomationElement kladblok = AutomationElement.RootElement.FindFirst(TreeScope.Descendants, new PropertyCondition(AutomationElement.NameProperty, "Kladblok"));
                Thread.Sleep(5000);

                InvokePattern invoke = kladblok.GetCurrentPattern(InvokePattern.Pattern) as InvokePattern;
                invoke.Invoke();
                Thread.Sleep(5000);


                // Read the file and retype it line by line.  
                string line;
                System.IO.StreamReader file = new System.IO.StreamReader(filelocation + boek);
                while ((line = file.ReadLine()) != null)
                {
                    Typ(line);
                }

                file.Close();

                //Find "Bestand" button

                AutomationElement notepadWindow = AutomationElement.RootElement.FindFirst(TreeScope.Descendants, new PropertyCondition(AutomationElement.ClassNameProperty, "Notepad"));
                Thread.Sleep(2000);

                AutomationElement filebutton = FindElementByNameAndControlType(notepadWindow, "Bestand", ControlType.MenuItem);

                Thread.Sleep(1000);

                //Expand "Bestand" menu

                ExpandCollapsePattern expandfilebutton = filebutton.GetCurrentPattern(ExpandCollapsePattern.Pattern) as ExpandCollapsePattern;
                expandfilebutton.Expand();
                Thread.Sleep(2000);

                //Find "Opslaan als" button

                AutomationElement savebutton = FindElementByNameAndControlType(notepadWindow, "Opslaan als...", ControlType.MenuItem);
                Thread.Sleep(1000);

                //Invoke "Opslaan als" button

                InvokePattern saveInvoke = savebutton.GetCurrentPattern(InvokePattern.Pattern) as InvokePattern;
                saveInvoke.Invoke();
                Thread.Sleep(1000);

                //Enter name of file

                Typ(boek);
                Thread.Sleep(2000);

                //Close notepad
                AutomationElement closebutton = FindElementByNameAndControlType(notepadWindow, "Sluiten", ControlType.Button);
                Thread.Sleep(2000);

                InvokePattern closeInvoke = closebutton.GetCurrentPattern(InvokePattern.Pattern) as InvokePattern;
                closeInvoke.Invoke();
                Thread.Sleep(2000);

                Console.WriteLine("Retyped " + boek);

            }

            Console.WriteLine("All books retyped, end of test");
        }

        static void Typ(string zin)
        {
            foreach (char c in zin)
            {
                if (c.Equals('('))
                {
                    SendKeys.SendWait("{(}");
                }else

                if (c.Equals(')'))
                {
                    SendKeys.SendWait("{)}");
                }else

                if (c.Equals('+'))
                {
                    SendKeys.SendWait("{+}");
                }else

                if (c.Equals('"'))
                {
                    SendKeys.SendWait(Char.ToString(c));
                    SendKeys.SendWait(" ");
                }else

                if (c.Equals('{'))
                {
                    SendKeys.SendWait("{{}");
                }else
                
                if (c.Equals('}'))
                {
                    SendKeys.SendWait("{}}");
                }
                else
                {
                    SendKeys.SendWait(Char.ToString(c));
                }

                Thread.Sleep(200);

            }
            SendKeys.SendWait("{ENTER}");
            Thread.Sleep(300);
        }

        //Method to find a UI element.
        //param root; AutomationElement representing the root of the search tree
        //param name; String containing name of the UI element to be found
        //param controltype; Object 

        static AutomationElement FindElementByNameAndControlType(AutomationElement root, string name, object controltype) 
        {
            AutomationElement element;

            var controlTypeProperty = new PropertyCondition(AutomationElement.ControlTypeProperty, controltype);
            var nameProperty = new PropertyCondition(AutomationElement.NameProperty, name);
            var andProperty = new AndCondition(controlTypeProperty, nameProperty);

            element = root.FindFirst(TreeScope.Descendants, andProperty);

            return element;
        }
    }
}
