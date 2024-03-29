﻿using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading;

namespace DocusignPlanners
{
    class Program
    {
        static void Main(string[] args)
        {
            //add reference to Microsoft Office Interop Excell dll
            string excelFileName = "Planners.xlsx";
            string soureFilePath = AppDomain.CurrentDomain.BaseDirectory + excelFileName;
            if (File.Exists(AppDomain.CurrentDomain.BaseDirectory + excelFileName))
            {
                Microsoft.Office.Interop.Excel.Application xlApplication = new Microsoft.Office.Interop.Excel.Application();
                // xlApplication.Visible = false;
                Microsoft.Office.Interop.Excel.Workbook workbook = xlApplication.Workbooks.Open(soureFilePath);
                foreach (Microsoft.Office.Interop.Excel.Worksheet worksheet in workbook.Worksheets)
                {
                    worksheet.Activate();
                    //workbook.SaveAs(AppDomain.CurrentDomain.BaseDirectory + worksheet.Name + ".txt", Microsoft.Office.Interop.Excel.XlFileFormat.xlCurrentPlatformText);
                    workbook.SaveAs(AppDomain.CurrentDomain.BaseDirectory + "Planners.txt", Microsoft.Office.Interop.Excel.XlFileFormat.xlCurrentPlatformText);

                    System.Runtime.InteropServices.Marshal.ReleaseComObject(worksheet);
                }//End foreach worksheet
                //Close and quit workbook and xlApplication
                workbook.Close();
                xlApplication.Quit();

                System.Runtime.InteropServices.Marshal.ReleaseComObject(workbook);
                System.Runtime.InteropServices.Marshal.ReleaseComObject(xlApplication);
                workbook = null;
                xlApplication = null;

                //Instantiate list from DataValues.cs 
                List<DataValues> dataValues = new List<DataValues>();
                //File name
                string fileName = "Planners.txt";

                //Check if 'Planners.txt' exists
                if (File.Exists(AppDomain.CurrentDomain.BaseDirectory + fileName))
                {
                    //Generate list from reading file
                    List<string> lines = File.ReadAllLines(AppDomain.CurrentDomain.BaseDirectory + fileName).ToList();

                    //Loop through each line and set variables
                    foreach (string line in lines)
                    {
                        string[] values = line.Split('\t'); //Tab delimited
                        char[] charsToTrim = { '"' }; //Trim double quotes

                        //Initialize variables 
                        string vendorPartner = values[1];
                        string contactFirstName;
                        string contactLastName;
                        string email;
                        string advertisingFund;
                        string baseMembership;
                        string proFixedCircular;
                        string proCustomizableCircular;
                        string proBOM;
                        string farmFixedCircular;
                        string farmCustomCircular;
                        string farmVOM;
                        string gardenMasterNationalCircular;
                        string gardenMasterVOM;
                        string proMonthlyBuys;
                        string protradeAds;
                        string proAchieversClub;
                        string totalBudget = values[26];
                        //string notes = values[27];

                        string notes = "";


                        //Logic for setting variables if data produces null or empty values
                        if (values[9] == "")
                        {
                            email = "NO_DATA";
                        }
                        else { email = values[9]; }

                        if (values[1] == "")
                        {
                            vendorPartner = "NO_DATA";
                            email = "NO_DATA";
                        }
                        else { vendorPartner = values[1]; }

                        if (totalBudget == "0")
                        {
                            email = "NO_DATA";
                            totalBudget = "NO_DATA";
                        }

                        if (values[7] == "")
                        {
                            contactFirstName = "NO_DATA";
                            email = "NO_DATA";
                        }
                        else { contactFirstName = values[7]; }

                        if (values[8] == "")
                        {
                            contactLastName = "NO_DATA";
                            email = "NO_DATA";
                        }
                        else { contactLastName = values[8]; }

                        if (values[10] == "")
                        { advertisingFund = "0"; }
                        else { advertisingFund = values[10]; }

                        if (values[12] == "")
                        { baseMembership = "0"; }
                        else { baseMembership = values[12]; }

                        if (values[13] == "")
                        { proFixedCircular = "0"; }
                        else { proFixedCircular = values[13]; }

                        if (values[14] == "")
                        { proCustomizableCircular = "0"; }
                        else { proCustomizableCircular = values[14]; }

                        if (values[15] == "")
                        { proBOM = "0"; }
                        else { proBOM = values[15]; }

                        if (values[19] == "")
                        { farmFixedCircular = "0"; }
                        else { farmFixedCircular = values[19]; }

                        if (values[20] == "")
                        { farmCustomCircular = "0"; }
                        else { farmCustomCircular = values[20]; }

                        if (values[21] == "")
                        { farmVOM = "0"; }
                        else { farmVOM = values[21]; }

                        if (values[22] == "")
                        { gardenMasterNationalCircular = "0"; }
                        else { gardenMasterNationalCircular = values[22]; }

                        if (values[23] == "")
                        { gardenMasterVOM = "0"; }
                        else { gardenMasterVOM = values[23]; }

                        if (values[11] == "")
                        { proMonthlyBuys = "0"; }
                        else { proMonthlyBuys = values[11]; }

                        if (values[17] == "")
                        { protradeAds = "0"; }
                        else { protradeAds = values[17]; }

                        if (values[16] == "")
                        { proAchieversClub = "0"; }
                        else { proAchieversClub = values[16]; }

                        //Trim double quotes and commas from data 
                        string trimmedAdvertisingFund = advertisingFund.Trim(charsToTrim).Replace(",", "");
                        string trimmedBaseMembership = baseMembership.Trim(charsToTrim).Replace(",", "");
                        string trimmedProFixedCircular = proFixedCircular.Trim(charsToTrim).Replace(",", "");
                        string trimmedProCustomizableCircular = proCustomizableCircular.Trim(charsToTrim).Replace(",", "");
                        string trimmedProBOM = proBOM.Trim(charsToTrim).Replace(",", "");
                        string trimmedFarmFixedCircular = farmFixedCircular.Trim(charsToTrim).Replace(",", "");
                        string trimmedFarmCustomCircular = farmCustomCircular.Trim(charsToTrim).Replace(",", "");
                        string trimmedFarmVOM = farmVOM.Trim(charsToTrim).Replace(",", "");
                        string trimmedGardenMasterNationalCircular = gardenMasterNationalCircular.Trim(charsToTrim).Replace(",", "");
                        string trimmedGardenMasterVOM = gardenMasterVOM.Trim(charsToTrim).Replace(",", "");
                        string trimmedProMonthlyBuys = proMonthlyBuys.Trim(charsToTrim).Replace(",", "");
                        string trimmedProTradeAds = protradeAds.Trim(charsToTrim).Replace(",", "");
                        string trimmedProAchieversClub = proAchieversClub.Trim(charsToTrim).Replace(",", "");
                        string trimmedNotes = notes.Trim(charsToTrim).Replace(",", "");
                        string trimmedTotalBudget = totalBudget.Trim(charsToTrim).Replace(",", "");
                        string trimmedVendorPartner = vendorPartner.Trim(charsToTrim).Replace(",", "");
                        string trimmedEmail = email.Trim(charsToTrim).Replace(",", "");
                        string trimmedContactFirstName = contactFirstName.Trim(charsToTrim).Replace(",", "");
                        string trimmedContactLastName = contactLastName.Trim(charsToTrim).Replace(",", "");

                        //Logic to bypass header
                        if (
                            values[0] != "ContactDisplayName" &&
                            trimmedAdvertisingFund != "AdvertisingFund" &&
                            trimmedBaseMembership != "BaseMembershipProgram" &&
                            trimmedProFixedCircular != "PROHardwareFixedCircular" &&
                            trimmedProCustomizableCircular != "PROHardwareCustomizableCircular" &&
                            trimmedProBOM != "PROHardwareBargainoftheMonth" &&
                            trimmedFarmFixedCircular != "FARMMARTFixedCircular" &&
                            trimmedFarmCustomCircular != "FARMMARTCustomizableCircular" &&
                            trimmedFarmVOM != "FARMMARTValueoftheMonth" &&
                            trimmedGardenMasterNationalCircular != "GardenMasterNationalCircular" &&
                            trimmedGardenMasterVOM != "GardenMasterValueoftheMonth" &&
                            trimmedProMonthlyBuys != "PROMonthlyBuys" &&
                            trimmedProTradeAds != "PROHardwareTradeAd" &&
                            trimmedProAchieversClub != "PROHardwareAchieversClub"
                        )
                        {
                            //Convert string variables to int for addition of data values
                            //Variables for Base Membership
                            int trimmedAdvertisingFundInt = int.Parse(trimmedAdvertisingFund);
                            int trimmedBaseMembershipInt = int.Parse(trimmedBaseMembership);
                            int finalBaseMembershipInt = trimmedAdvertisingFundInt + trimmedBaseMembershipInt;
                            string finalBaseMembership = finalBaseMembershipInt.ToString();
                            //Variables for PRO Print Circulars
                            int trimmedProFixedCircularInt = int.Parse(trimmedProFixedCircular);
                            int trimmedProCustomizableCircularInt = int.Parse(trimmedProCustomizableCircular);
                            int trimmedProBOMInt = int.Parse(trimmedProBOM);
                            int trimmedFarmFixedCircularInt = int.Parse(trimmedFarmFixedCircular);
                            int trimmedFarmCustomCircularInt = int.Parse(trimmedFarmCustomCircular);
                            int trimmedFarmVOMInt = int.Parse(trimmedFarmVOM);
                            int trimmedGardenMasterNationalCircularInt = int.Parse(trimmedGardenMasterNationalCircular);
                            int trimmedGardenMasterVOMInt = int.Parse(trimmedGardenMasterVOM);
                            int finalPrintCircularInt =
                                trimmedProFixedCircularInt +
                                trimmedProCustomizableCircularInt +
                                trimmedProBOMInt +
                                trimmedFarmFixedCircularInt +
                                trimmedFarmCustomCircularInt +
                                trimmedFarmVOMInt +
                                trimmedGardenMasterNationalCircularInt +
                                trimmedGardenMasterVOMInt
                            ;
                            string finalPrintCircular = finalPrintCircularInt.ToString();

                            //Variables for Monthly Buys
                            int trimmedProMonthlyBuysInt = int.Parse(trimmedProMonthlyBuys);
                            string finalProMonthlyBuys = trimmedProMonthlyBuysInt.ToString();

                            //Variables for Trade Advertisements
                            int trimmedProTradeAdsInt = int.Parse(trimmedProTradeAds);
                            string finalProtradeAds = trimmedProTradeAdsInt.ToString();

                            //Variables for Achievers Club
                            int trimmedProAchieversClubInt = int.Parse(trimmedProAchieversClub);
                            string finalProAchieversClub = trimmedProAchieversClubInt.ToString();

                            //Get the total
                            int trimmedTotalInt = finalBaseMembershipInt + finalPrintCircularInt + trimmedProMonthlyBuysInt + trimmedProTradeAdsInt + trimmedProAchieversClubInt;
                            string finalTrimmedTotal = trimmedTotalInt.ToString();
                            if (finalTrimmedTotal == "0")
                            {
                                finalTrimmedTotal = "NO_DATA";
                            }
                            //Initialize DataValues list
                            DataValues newDataValues = new DataValues();

                            //Get data for DataValues list
                            newDataValues.Email = trimmedEmail;
                            newDataValues.VendorPartner = trimmedVendorPartner;
                            newDataValues.BaseMembership = finalBaseMembership;
                            newDataValues.PrintCircular = finalPrintCircular;
                            newDataValues.ProMonthlyBuys = finalProMonthlyBuys;
                            newDataValues.TradeAdvertisments = finalProtradeAds;
                            newDataValues.AchieversClubSponsorship = finalProAchieversClub;
                            newDataValues.TotalPromotionalBudget = finalTrimmedTotal;
                            newDataValues.TotalPromotionalBudgetSmall = finalTrimmedTotal;
                            newDataValues.Notes = trimmedNotes;
                            newDataValues.ContactName = trimmedContactFirstName + " " + trimmedContactLastName;

                            //Set data for DataValues list 
                            dataValues.Add(newDataValues);
                        }
                    }
                    //Instantiate new list for csv file
                    List<string> output = new List<string>();

                    //Instantiate new list for errors
                    List<string> errors = new List<string>();

                    //Set error file status
                    bool errorFileStatus = false;

                    //Initialize header for error file
                    errors.Add("Name,Email,VendorPartner,BaseMembership,PrintCircular,ProMonthlyBuys,TradeAdvertisments,AchieversClubSponsorship,TotalPromotionalBudget,Notes,TotalPromotionalBudgetSmall,ContactName");

                    //Initialize header for csv
                    output.Add("Sender::Name,Sender::Email,Sender::VendorPartner,Sender::BaseMembership,Sender::PrintCircular,Sender::ProMonthlyBuys,Sender::TradeAdvertisments,Sender::AchieversClubSponsorship,Sender::TotalPromotionalBudget,Sender::Notes,Sender::TotalPromotionalBudgetSmall,Sender::ContactName");

                    //Loop and add data for each line of new csv file
                    foreach (var dataValue in dataValues)
                    {   //If file has missing data, remove from CSV and create error file and set error file status to true.
                        if (dataValue.ContactName == "NO_DATA NO_DATA" || dataValue.Email == "NO_DATA" || dataValue.TotalPromotionalBudget == "NO_DATA" || dataValue.TotalPromotionalBudgetSmall == "NO_DATA")
                        {
                            output.Remove($"{ dataValue.ContactName},{ dataValue.Email},{ dataValue.VendorPartner},{ dataValue.BaseMembership},{ dataValue.PrintCircular},{ dataValue.ProMonthlyBuys},{ dataValue.TradeAdvertisments},{ dataValue.AchieversClubSponsorship},{ dataValue.TotalPromotionalBudget},{ dataValue.Notes},{ dataValue.TotalPromotionalBudgetSmall},{dataValue.ContactName}");
                            errors.Add($"{ dataValue.ContactName},{ dataValue.Email},{ dataValue.VendorPartner},{ dataValue.BaseMembership},{ dataValue.PrintCircular},{ dataValue.ProMonthlyBuys},{ dataValue.TradeAdvertisments},{ dataValue.AchieversClubSponsorship},{ dataValue.TotalPromotionalBudget},{ dataValue.Notes},{ dataValue.TotalPromotionalBudgetSmall},{dataValue.ContactName}");
                            errorFileStatus = true;
                        }
                        else
                        {   //If file has no errors, create CSV
                            output.Add($"{ dataValue.ContactName},{ dataValue.Email},{ dataValue.VendorPartner},{ dataValue.BaseMembership},{ dataValue.PrintCircular},{ dataValue.ProMonthlyBuys},{ dataValue.TradeAdvertisments},{ dataValue.AchieversClubSponsorship},{ dataValue.TotalPromotionalBudget},{ dataValue.Notes},{ dataValue.TotalPromotionalBudgetSmall},{dataValue.ContactName}");
                        }
                    }

                    //Create unique file name using current date time down to the minute
                    DateTime currentTime = DateTime.Now;
                    string now = currentTime.ToString("MM_dd_yyyy_HHmmss");
                    string docusignFileName = "Docusign_Upload";
                    string dotCsv = ".csv";
                    string docusignErrorFileName = "Docusign_Errors";

                    //If error file status is true, write error file and write cleaned CSV file.
                    if (errorFileStatus == true)
                    {
                        File.WriteAllLines(AppDomain.CurrentDomain.BaseDirectory + docusignErrorFileName + "_" + now + dotCsv, errors);
                        File.WriteAllLines(AppDomain.CurrentDomain.BaseDirectory + docusignFileName + "_" + now + dotCsv, output);
                        Console.WriteLine("Docusign CSV File created successfully! ");
                        Console.WriteLine();
                        Console.WriteLine("Your formatted Docusign CSV file is in this location:");
                        Console.WriteLine(Directory.GetCurrentDirectory());
                        Console.WriteLine();
                        Console.WriteLine("***Errors detected*** \r\n See Docusign_Error file for details.");
                        Console.WriteLine();
                        //File.Delete(fileName);
                        Console.WriteLine("You are now being directed to Docusign.com");
                        File.Delete(fileName);
                        System.Diagnostics.Process.Start("https://account.docusign.com/");
                        Thread.Sleep(3000);
                        System.Environment.Exit(0);
                        //Console.ReadLine();
                    }
                    else
                    {   //If no errors, write CSV file.
                        File.WriteAllLines(AppDomain.CurrentDomain.BaseDirectory + docusignFileName + "_" + now + dotCsv, output);
                        Console.WriteLine("Docusign CSV File created successfully! ");
                        Console.WriteLine();
                        Console.WriteLine("Your formatted Docusign CSV file is in this location:");
                        Console.WriteLine(Directory.GetCurrentDirectory());
                        Console.WriteLine();
                        //File.Delete(fileName);
                        Console.WriteLine("You are now being directed to Docusign.com");
                        File.Delete(fileName);
                        System.Diagnostics.Process.Start("https://account.docusign.com/");
                        Thread.Sleep(3000);
                        System.Environment.Exit(0);
                        //Console.ReadLine();

                    }

                }//End if for fileName
                else
                {
                    //If Planners.txt does not exist, raise error. 
                    Console.WriteLine("Excel file 'Planners.txt' does not exist in the current directory!");
                    Console.ReadLine();
                }
            }
            else
            {   //If Planners.txt does not exist, raise error. 
                Console.WriteLine("Excel file 'Planners.xlsx' does not exist in the current directory!");
                Console.ReadLine();
            }

            //End old docusignplanners code/////////////////////////////////
        }//End Main()
    }//End class program
}//End namespace

