using System;
using System.Collections.Generic;
using System.IO;
using System.Net.Mime;
using System.Threading;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using OpenQA.Selenium;
using OpenQA.Selenium.Chrome;
using OpenQA.Selenium.Support.UI;
using OpenQA.Selenium.Interactions;
using Excel = Microsoft.Office.Interop.Excel;
using  System.Runtime.InteropServices;

namespace HelloWPF
{
    public class AddRentalUnits
    {

        public void CreateRental(string file, string name)
            {
                var optionOne = new ChromeOptions();
                optionOne.AddArgument("test-type");
                IWebDriver driver = new ChromeDriver(optionOne);

                driver.Navigate().GoToUrl("https://dev-productadmin.vacationroost.com/"); // Dev


                //driver.Navigate().GoToUrl("https://productadmin.vacationroost.com");

                var wait = new WebDriverWait(driver, TimeSpan.FromSeconds(250));

                //if (driver.FindElement(By.XPath("//*[@id='ctl00_body_ProductsTaskSelection_LodgingTasksGroup']/legend")).Displayed)
                //{
                //    Thread.Sleep(500);
                driver.Manage().Timeouts().ImplicitlyWait(TimeSpan.FromSeconds(145));
                IWebElement editSupplier = wait.Until(x => x.FindElement(By.XPath("//a[contains(.,'Edit a Lodging Supplier')]")));
                editSupplier.Click();   
            //}

                string fileName = file;
                string suppName = name;

                IWebElement supplierName = wait.Until(x => x.FindElement(By.Id("ctl00_body_EntitySelection1_SupplierNameField")));
                supplierName.SendKeys(name);

                IWebElement supplierSearchButton = wait.Until(x => x.FindElement(By.Id("ctl00_body_EntitySelection1_EntitySelectionSearchButton"))); 
                supplierSearchButton.Click();

                IWebElement selectLinkForSupplier = wait.Until(x => x.FindElement(By.Id("ctl00_body_EntitySelection1_SupplierLodgingRepeater_ctl00_EditALodgingSupplierSelectionLineItem_SelectButton"))); 
                selectLinkForSupplier.Click();

                IWebElement createStandAloneRentalUnit = wait.Until(x => x.FindElement(By.Id("ctl00_body_DashboardControl1_TaskSelection_ctl01_RentalUnitTaskGroup_RentalUnitTasksControl_CreateRentalUnitSubTaskGroup_CreateARentalUnitTaskLink")));
                createStandAloneRentalUnit.Click();

                var rentalUnitInfos = new List<RentalUnitInformation>();
                var finalUrl = "";
                var finalComplex = "";

            if (File.Exists(fileName))
            {
                var xlApp = new Excel.Application();
                var xlWorkBook = xlApp.Workbooks.Open(file, false);
                var xlWorksheet = (Excel.Worksheet) xlWorkBook.Worksheets.Item[1];

                Excel.Range credentials = xlWorksheet.UsedRange;

                int rowCount = credentials.Rows.Count;

                try
                {
                    for (int i = 1; i < rowCount; i++)
                    {
                        string building = null;
                        string rooms = null;
                        string numberBath = null;
                        string numberGuess = null;

                        var complex = Convert.ToString(credentials.Cells[i + 1, 1].Value2);
                        var description = credentials.Cells[i + 1, 2].Value2.ToString();
                        building = credentials.Cells[i + 1, 3].Value2.ToString();
                        rooms = credentials.Cells[i + 1, 4].Value2.ToString();
                        var uNumber = credentials.Cells[i + 1, 5].Value2.ToString();
                        var uDescription = credentials.Cells[i + 1, 6].Value2.ToString();
                        numberBath = credentials.Cells[i + 1, 7].Value2.ToString();
                        numberGuess = credentials.Cells[i + 1, 8].Value2.ToString();

                        // Unit Name
                        var waiting = new WebDriverWait(driver, TimeSpan.FromSeconds(260));

                        IWebElement complexName =
                            waiting.Until(x => x.FindElement(By.Id("ctl00_body_InputForm_ComplexNameTextbox")));
                        complexName.SendKeys(complex);

                        // Complex Description

                        IWebElement iframe =
                            driver.FindElement(
                                By.XPath(
                                    "//*[@id='ctl00_body_InputForm_ComplexDescriptionLabel']/span[1]/div/div[2]/iframe"));
                        driver.SwitchTo().Frame(iframe);

                        IWebElement complexDescription = waiting.Until(x => x.FindElement(By.TagName("body")));

                        complexDescription.SendKeys(description);

                        driver.SwitchTo().DefaultContent();


                        // Building Type

                        switch (building)
                        {
                            case "Apartment":
                                IWebElement apartmentbuilding =
                                    waiting.Until(x => x.FindElement(By.Id("ctl00_body_InputForm_BuildingTypeList")));
                                var selectedApartment = new SelectElement(apartmentbuilding);
                                selectedApartment.SelectByText("Apartment");
                                break;

                            case "Bed and Breakfast":
                                IWebElement bedAndBreakfastBuilding =
                                    waiting.Until(x => x.FindElement(By.Id("ctl00_body_InputForm_BuildingTypeList")));
                                var selectedBedAnBreakfast = new SelectElement(bedAndBreakfastBuilding);
                                selectedBedAnBreakfast.SelectByText("Bed & Breakfast");
                                break;

                            case "Boutique Hotel":
                                IWebElement boutiqueBuilding =
                                    waiting.Until(x => x.FindElement(By.Id("ctl00_body_InputForm_BuildingTypeList")));
                                var selectedboutique = new SelectElement(boutiqueBuilding);
                                selectedboutique.SelectByText("Hotel Inn");
                                break;

                            case "Bungalow":
                                IWebElement bungalow =
                                    waiting.Until(x => x.FindElement(By.Id("ctl00_body_InputForm_BuildingTypeList")));
                                var selectedBungalow = new SelectElement(bungalow);
                                selectedBungalow.SelectByText("Bungalow");
                                break;

                            case "Camping Ground":
                                IWebElement campingGround =
                                    waiting.Until(x => x.FindElement(By.Id("ctl00_body_InputForm_BuildingTypeList")));
                                var selectedCampingGround = new SelectElement(campingGround);
                                selectedCampingGround.SelectByText("Camping Ground");
                                break;

                            case "Chalet":
                                IWebElement chaletBuilding =
                                    waiting.Until(x => x.FindElement(By.Id("ctl00_body_InputForm_BuildingTypeList")));
                                var selectedchalet = new SelectElement(chaletBuilding);
                                selectedchalet.SelectByText("Chalet");
                                break;

                            case "Cottage":
                                IWebElement cottageBuilding =
                                    waiting.Until(x => x.FindElement(By.Id("ctl00_body_InputForm_BuildingTypeList")));
                                var selectedCottage = new SelectElement(cottageBuilding);
                                selectedCottage.SelectByText("Cabin / Lodge / Cottage");
                                break;

                            case "Farm":
                                IWebElement farm =
                                    waiting.Until(x => x.FindElement(By.Id("ctl00_body_InputForm_BuildingTypeList")));
                                var selectedFarm = new SelectElement(farm);
                                selectedFarm.SelectByText("Farm");
                                break;

                            case "Guesthouse":
                                IWebElement guestHouse =
                                    waiting.Until(x => x.FindElement(By.Id("ctl00_body_InputForm_BuildingTypeList")));
                                var selectedGuestHouse = new SelectElement(guestHouse);
                                selectedGuestHouse.SelectByText("Guest House");
                                break;

                            case "Home":
                                IWebElement homeBuilding =
                                    waiting.Until(x => x.FindElement(By.Id("ctl00_body_InputForm_BuildingTypeList")));
                                var selectedHome = new SelectElement(homeBuilding);
                                selectedHome.SelectByText("Vacation Home");
                                break;

                            case "Homestay":
                                IWebElement homestayBuilding =
                                    waiting.Until(x => x.FindElement(By.Id("ctl00_body_InputForm_BuildingTypeList")));
                                var selectedHomestay = new SelectElement(homestayBuilding);
                                selectedHomestay.SelectByText("Vacation Home");
                                break;

                            case "Luxury Yacht":
                                IWebElement luxuryYacht =
                                    waiting.Until(x => x.FindElement(By.Id("ctl00_body_InputForm_BuildingTypeList")));
                                var selectedLuxuryYacht = new SelectElement(luxuryYacht);
                                selectedLuxuryYacht.SelectByText("Yacht");
                                break;

                            case "Resort":
                                IWebElement resortBuilding =
                                    waiting.Until(x => x.FindElement(By.Id("ctl00_body_InputForm_BuildingTypeList")));
                                var selectedResort = new SelectElement(resortBuilding);
                                selectedResort.SelectByText("All Inclusive Hotel");
                                break;

                            case "Serviced Apartment":
                                IWebElement serviceBuilding =
                                    waiting.Until(x => x.FindElement(By.Id("ctl00_body_InputForm_BuildingTypeList")));
                                var selectedServiceAppt = new SelectElement(serviceBuilding);
                                selectedServiceAppt.SelectByText("Apartment");
                                break;

                            case "Villa":
                                IWebElement villaBuilding =
                                    waiting.Until(x => x.FindElement(By.Id("ctl00_body_InputForm_BuildingTypeList")));
                                var selectedVilla = new SelectElement(villaBuilding);
                                selectedVilla.SelectByText("Villa");
                                break;
                        } // End Building Type

                        Thread.Sleep(500);
                        // Floor Plan Type

                        switch (building)
                        {

                            case "Apartment":
                                IWebElement apartmentType =
                                    waiting.Until(x => x.FindElement(By.Id("FloorPlanTypeList")));
                                var act1 = new Actions(driver);
                                act1.MoveToElement(apartmentType)
                                    .Click(driver.FindElement(By.XPath("//*[@id='FloorPlanTypeList']")))
                                    .Build()
                                    .Perform();
                                Thread.Sleep(800);
                                var selectedFloorTypeAppartment = new SelectElement(apartmentType);
                                selectedFloorTypeAppartment.SelectByText("Apartment");
                                break;

                            case "Bed and Breakfast":
                                IWebElement bedAndBreakfastType =
                                    waiting.Until(x => x.FindElement(By.Id("FloorPlanTypeList")));
                                var act2 = new Actions(driver);
                                act2.MoveToElement(bedAndBreakfastType)
                                    .Click(driver.FindElement(By.XPath("//*[@id='FloorPlanTypeList']")))
                                    .Build()
                                    .Perform();
                                Thread.Sleep(800);
                                var selectedFloorTypeBedAndBreakfast = new SelectElement(bedAndBreakfastType);
                                selectedFloorTypeBedAndBreakfast.SelectByText("Home");
                                break;

                            case "Boutique Hotel":
                                IWebElement boutiqueType =
                                    waiting.Until(x => x.FindElement(By.Id("FloorPlanTypeList")));
                                var act3 = new Actions(driver);
                                act3.MoveToElement(boutiqueType)
                                    .Click(driver.FindElement(By.XPath("//*[@id='FloorPlanTypeList']")))
                                    .Build()
                                    .Perform();
                                Thread.Sleep(800);
                                var selectedFloorTypeBoutique = new SelectElement(boutiqueType);
                                selectedFloorTypeBoutique.SelectByText("Condo");
                                break;

                            case "Bungalow":
                                IWebElement bungalow = waiting.Until(x => x.FindElement(By.Id("FloorPlanTypeList")));
                                var act4 = new Actions(driver);
                                act4.MoveToElement(bungalow)
                                    .Click(driver.FindElement(By.XPath("//*[@id='FloorPlanTypeList']")))
                                    .Build()
                                    .Perform();
                                Thread.Sleep(800);
                                var selectedBungalow = new SelectElement(bungalow);
                                selectedBungalow.SelectByText("Home");
                                break;

                            case "Camping Ground":
                                IWebElement campingGroundType =
                                    waiting.Until(x => x.FindElement(By.Id("FloorPlanTypeList")));
                                var act5 = new Actions(driver);
                                act5.MoveToElement(campingGroundType)
                                    .Click(driver.FindElement(By.XPath("//*[@id='FloorPlanTypeList']")))
                                    .Build()
                                    .Perform();
                                Thread.Sleep(800);
                                var selectedFloorCampingGround = new SelectElement(campingGroundType);
                                selectedFloorCampingGround.SelectByText("Camping Ground");
                                break;

                            case "Chalet":
                                IWebElement chaletType =
                                    waiting.Until(x => x.FindElement(By.Id("FloorPlanTypeList")));
                                var act6 = new Actions(driver);
                                act6.MoveToElement(chaletType)
                                    .Click(driver.FindElement(By.XPath("//*[@id='FloorPlanTypeList']")))
                                    .Build()
                                    .Perform();
                                Thread.Sleep(800);
                                var selectedFloorTypeChalet = new SelectElement(chaletType);
                                selectedFloorTypeChalet.SelectByText("Condo");
                                break;

                            case "Cottage":
                                IWebElement cottageType =
                                    waiting.Until(x => x.FindElement(By.Id("FloorPlanTypeList")));
                                var act7 = new Actions(driver);
                                act7.MoveToElement(cottageType)
                                    .Click(driver.FindElement(By.XPath("//*[@id='FloorPlanTypeList']")))
                                    .Build()
                                    .Perform();
                                Thread.Sleep(800);
                                var selectedFloorTypeCottage = new SelectElement(cottageType);
                                selectedFloorTypeCottage.SelectByText("Home");
                                break;

                            case "Farm":
                                IWebElement farmType = waiting.Until(x => x.FindElement(By.Id("FloorPlanTypeList")));
                                var act8 = new Actions(driver);
                                act8.MoveToElement(farmType)
                                    .Click(driver.FindElement(By.XPath("//*[@id='FloorPlanTypeList']")))
                                    .Build()
                                    .Perform();
                                Thread.Sleep(800);
                                var selectedFloorFarm = new SelectElement(farmType);
                                selectedFloorFarm.SelectByText("Home");
                                break;

                            case "Guesthouse":
                                IWebElement guestHouseType =
                                    waiting.Until(x => x.FindElement(By.Id("FloorPlanTypeList")));
                                var act9 = new Actions(driver);
                                act9.MoveToElement(guestHouseType)
                                    .Click(driver.FindElement(By.XPath("//*[@id='FloorPlanTypeList']")))
                                    .Build()
                                    .Perform();
                                Thread.Sleep(800);
                                var selectedFloorGuestHouse = new SelectElement(guestHouseType);
                                selectedFloorGuestHouse.SelectByText("Home");
                                break;

                            case "Home":
                                IWebElement homeType = waiting.Until(x => x.FindElement(By.Id("FloorPlanTypeList")));
                                var act10 = new Actions(driver);
                                act10.MoveToElement(homeType)
                                    .Click(driver.FindElement(By.XPath("//*[@id='FloorPlanTypeList']")))
                                    .Build()
                                    .Perform();
                                Thread.Sleep(800);
                                var selectedFloorTypeHome = new SelectElement(homeType);
                                selectedFloorTypeHome.SelectByText("Home");
                                break;

                            case "Homestay":
                                IWebElement homestayType =
                                    waiting.Until(x => x.FindElement(By.Id("FloorPlanTypeList")));
                                var act11 = new Actions(driver);
                                act11.MoveToElement(homestayType)
                                    .Click(driver.FindElement(By.XPath("//*[@id='FloorPlanTypeList']")))
                                    .Build()
                                    .Perform();
                                Thread.Sleep(800);
                                var selectedFloorTypeHomestay = new SelectElement(homestayType);
                                selectedFloorTypeHomestay.SelectByText("Home");
                                break;

                            case "Luxury Yacht":
                                IWebElement luxuryYachtType =
                                    waiting.Until(x => x.FindElement(By.Id("FloorPlanTypeList")));
                                var act12 = new Actions(driver);
                                act12.MoveToElement(luxuryYachtType)
                                    .Click(driver.FindElement(By.XPath("//*[@id='FloorPlanTypeList']")))
                                    .Build()
                                    .Perform();
                                Thread.Sleep(800);
                                var selectedFloorTypeLuxuryYacht = new SelectElement(luxuryYachtType);
                                selectedFloorTypeLuxuryYacht.SelectByText("Yacht");
                                break;

                            case "Resort":
                                IWebElement resortType =
                                    waiting.Until(x => x.FindElement(By.Id("FloorPlanTypeList")));
                                var act13 = new Actions(driver);
                                act13.MoveToElement(resortType)
                                    .Click(driver.FindElement(By.XPath("//*[@id='FloorPlanTypeList']")))
                                    .Build()
                                    .Perform();
                                Thread.Sleep(800);
                                var selectedFloorTypeResort = new SelectElement(resortType);
                                selectedFloorTypeResort.SelectByText("Condo");
                                break;

                            case "Serviced Apartment":
                                IWebElement serviceApptType =
                                    waiting.Until(x => x.FindElement(By.Id("FloorPlanTypeList")));
                                var act14 = new Actions(driver);
                                act14.MoveToElement(serviceApptType)
                                    .Click(driver.FindElement(By.XPath("//*[@id='FloorPlanTypeList']")))
                                    .Build()
                                    .Perform();
                                Thread.Sleep(800);
                                var selectedFloorTypeServiceAppt = new SelectElement(serviceApptType);
                                selectedFloorTypeServiceAppt.SelectByText("Apartment");
                                break;

                            case "Villa":
                                IWebElement villaType = waiting.Until(x => x.FindElement(By.Id("FloorPlanTypeList")));
                                var act15 = new Actions(driver);
                                act15.MoveToElement(villaType)
                                    .Click(driver.FindElement(By.XPath("//*[@id='FloorPlanTypeList']")))
                                    .Build()
                                    .Perform();
                                Thread.Sleep(800);
                                var selectedFloorTypeVilla = new SelectElement(villaType);
                                selectedFloorTypeVilla.SelectByText("Condo");
                                break;

                        } // End Floor Plan Type


                        Thread.Sleep(500);
                        // Bedrooms


                        switch (rooms)
                        {


                            case "1":

                                IWebElement oneBedroom = waiting.Until(x => x.FindElement(By.Id("BedroomsDropDown")));
                                var act = new Actions(driver);
                                act.MoveToElement(oneBedroom)
                                    .Click(driver.FindElement(By.XPath("//*[@id='BedroomsDropDown']")))
                                    .Build()
                                    .Perform();
                                Thread.Sleep(1000);
                                var selectedBedroomOne = new SelectElement(oneBedroom);
                                selectedBedroomOne.SelectByValue("1BR");
                                break;

                            case "2":

                                IWebElement twoBedroom = waiting.Until(x => x.FindElement(By.Id("BedroomsDropDown")));
                                var act2 = new Actions(driver);
                                act2.MoveToElement(twoBedroom)
                                    .Click(driver.FindElement(By.XPath("//*[@id='BedroomsDropDown']")))
                                    .Build()
                                    .Perform();
                                Thread.Sleep(1000);
                                var selectedBedroomTwo = new SelectElement(twoBedroom);
                                selectedBedroomTwo.SelectByValue("2BR");
                                break;

                            case "3":
                                IWebElement threeBedroom =
                                    waiting.Until(x => x.FindElement(By.Id("BedroomsDropDown")));
                                var act3 = new Actions(driver);
                                act3.MoveToElement(threeBedroom)
                                    .Click(driver.FindElement(By.XPath("//*[@id='BedroomsDropDown']")))
                                    .Build()
                                    .Perform();
                                Thread.Sleep(1000);
                                var selectedBedroomThree = new SelectElement(threeBedroom);
                                selectedBedroomThree.SelectByValue("3BR");
                                break;

                            case "4":
                                IWebElement fourBedrooms =
                                    waiting.Until(x => x.FindElement(By.Id("BedroomsDropDown")));
                                var act4 = new Actions(driver);
                                act4.MoveToElement(fourBedrooms)
                                    .Click(driver.FindElement(By.XPath("//*[@id='BedroomsDropDown']")))
                                    .Build()
                                    .Perform();
                                Thread.Sleep(1000);
                                var selectedBedroomFour = new SelectElement(fourBedrooms);
                                selectedBedroomFour.SelectByValue("4BR");
                                break;

                            case "5":
                                IWebElement fiveBedrooms =
                                    waiting.Until(x => x.FindElement(By.Id("BedroomsDropDown")));
                                var act5 = new Actions(driver);
                                act5.MoveToElement(fiveBedrooms)
                                    .Click(driver.FindElement(By.XPath("//*[@id='BedroomsDropDown']")))
                                    .Build()
                                    .Perform();
                                Thread.Sleep(1000);
                                var selectedBedroomFive = new SelectElement(fiveBedrooms);
                                selectedBedroomFive.SelectByValue("5BR");
                                break;

                            case "6":
                                IWebElement sixBedrooms =
                                    waiting.Until(x => x.FindElement(By.Id("BedroomsDropDown")));
                                var act6 = new Actions(driver);
                                act6.MoveToElement(sixBedrooms)
                                    .Click(driver.FindElement(By.XPath("//*[@id='BedroomsDropDown']")))
                                    .Build()
                                    .Perform();
                                Thread.Sleep(1000);
                                var selectedBedroomSix = new SelectElement(sixBedrooms);
                                selectedBedroomSix.SelectByValue("6BR");
                                break;

                            case "7":
                                IWebElement sevenBedrooms =
                                    waiting.Until(x => x.FindElement(By.Id("BedroomsDropDown")));
                                var act7 = new Actions(driver);
                                act7.MoveToElement(sevenBedrooms)
                                    .Click(driver.FindElement(By.XPath("//*[@id='BedroomsDropDown']")))
                                    .Build()
                                    .Perform();
                                Thread.Sleep(1000);
                                var selectedBedroomSeven = new SelectElement(sevenBedrooms);
                                selectedBedroomSeven.SelectByValue("7BR");
                                break;

                            case "8":
                                IWebElement eigthBedrooms =
                                    waiting.Until(x => x.FindElement(By.Id("BedroomsDropDown")));
                                var act8 = new Actions(driver);
                                act8.MoveToElement(eigthBedrooms)
                                    .Click(driver.FindElement(By.XPath("//*[@id='BedroomsDropDown']")))
                                    .Build()
                                    .Perform();
                                Thread.Sleep(1000);
                                var selectedBedroomEight = new SelectElement(eigthBedrooms);
                                selectedBedroomEight.SelectByValue("8BR");
                                break;

                            case "9":
                                IWebElement nineBedrooms =
                                    waiting.Until(x => x.FindElement(By.Id("BedroomsDropDown")));
                                var act9 = new Actions(driver);
                                act9.MoveToElement(nineBedrooms)
                                    .Click(driver.FindElement(By.XPath("//*[@id='BedroomsDropDown']")))
                                    .Build()
                                    .Perform();
                                Thread.Sleep(1000);
                                var selectedBedroomNine = new SelectElement(nineBedrooms);
                                selectedBedroomNine.SelectByText("9BR");
                                break;

                            case "10":
                                IWebElement tenBedrooms =
                                    waiting.Until(x => x.FindElement(By.Id("BedroomsDropDown")));
                                var act10 = new Actions(driver);
                                act10.MoveToElement(tenBedrooms)
                                    .Click(driver.FindElement(By.XPath("//*[@id='BedroomsDropDown']")))
                                    .Build()
                                    .Perform();
                                Thread.Sleep(1000);
                                var selectedBedroomTen = new SelectElement(tenBedrooms);
                                selectedBedroomTen.SelectByText("10BR");
                                break;

                            case "11":
                                IWebElement elevenBedrooms =
                                    waiting.Until(x => x.FindElement(By.Id("BedroomsDropDown")));
                                var act11 = new Actions(driver);
                                act11.MoveToElement(elevenBedrooms)
                                    .Click(driver.FindElement(By.XPath("//*[@id='BedroomsDropDown']")))
                                    .Build()
                                    .Perform();
                                Thread.Sleep(1000);
                                var selectedBedroomEleven = new SelectElement(elevenBedrooms);
                                selectedBedroomEleven.SelectByText("11BR");
                                break;

                            case "12":
                                IWebElement twelveBedrooms =
                                    waiting.Until(x => x.FindElement(By.Id("BedroomsDropDown")));
                                var act12 = new Actions(driver);
                                act12.MoveToElement(twelveBedrooms)
                                    .Click(driver.FindElement(By.XPath("//*[@id='BedroomsDropDown']")))
                                    .Build()
                                    .Perform();
                                Thread.Sleep(1000);
                                var selectedBedroomTwelve = new SelectElement(twelveBedrooms);
                                selectedBedroomTwelve.SelectByText("12BR");
                                break;

                            case "13":
                                IWebElement thirteenBedrooms =
                                    waiting.Until(x => x.FindElement(By.Id("BedroomsDropDown")));
                                var act13 = new Actions(driver);
                                act13.MoveToElement(thirteenBedrooms)
                                    .Click(driver.FindElement(By.XPath("//*[@id='BedroomsDropDown']")))
                                    .Build()
                                    .Perform();
                                Thread.Sleep(1000);
                                var selectedBedroomThirteen = new SelectElement(thirteenBedrooms);
                                selectedBedroomThirteen.SelectByText("13BR");
                                break;

                            case "14":
                                IWebElement fourteenBedrooms =
                                    waiting.Until(x => x.FindElement(By.Id("BedroomsDropDown")));
                                var act14 = new Actions(driver);
                                act14.MoveToElement(fourteenBedrooms)
                                    .Click(driver.FindElement(By.XPath("//*[@id='BedroomsDropDown']")))
                                    .Build()
                                    .Perform();
                                Thread.Sleep(1000);
                                var selectedBedroomFourteen = new SelectElement(fourteenBedrooms);
                                selectedBedroomFourteen.SelectByText("14BR");
                                break;

                            case "15":
                                IWebElement fifteenBedrooms =
                                    waiting.Until(x => x.FindElement(By.Id("BedroomsDropDown")));
                                var act15 = new Actions(driver);
                                act15.MoveToElement(fifteenBedrooms)
                                    .Click(driver.FindElement(By.XPath("//*[@id='BedroomsDropDown']")))
                                    .Build()
                                    .Perform();
                                Thread.Sleep(1000);
                                var selectedBedroomFifteen = new SelectElement(fifteenBedrooms);
                                selectedBedroomFifteen.SelectByText("15BR");
                                break;

                            case "16":
                                IWebElement sixteenBedrooms =
                                    waiting.Until(x => x.FindElement(By.Id("BedroomsDropDown")));
                                var act16 = new Actions(driver);
                                act16.MoveToElement(sixteenBedrooms)
                                    .Click(driver.FindElement(By.XPath("//*[@id='BedroomsDropDown']")))
                                    .Build()
                                    .Perform();
                                Thread.Sleep(1000);
                                var selectedBedroomSixteen = new SelectElement(sixteenBedrooms);
                                selectedBedroomSixteen.SelectByText("16BR");
                                break;

                            case "17":
                                IWebElement seventeenBedrooms =
                                    waiting.Until(x => x.FindElement(By.Id("BedroomsDropDown")));
                                var act17 = new Actions(driver);
                                act17.MoveToElement(seventeenBedrooms)
                                    .Click(driver.FindElement(By.XPath("//*[@id='BedroomsDropDown']")))
                                    .Build()
                                    .Perform();
                                Thread.Sleep(1000);
                                var selectedBedroomSeventeen = new SelectElement(seventeenBedrooms);
                                selectedBedroomSeventeen.SelectByText("17BR");
                                break;

                            case "18":
                                IWebElement eighteenBedrooms =
                                    waiting.Until(x => x.FindElement(By.Id("BedroomsDropDown")));
                                var act18 = new Actions(driver);
                                act18.MoveToElement(eighteenBedrooms)
                                    .Click(driver.FindElement(By.XPath("//*[@id='BedroomsDropDown']")))
                                    .Build()
                                    .Perform();
                                Thread.Sleep(1000);
                                var selectedBedroomEighteen = new SelectElement(eighteenBedrooms);
                                selectedBedroomEighteen.SelectByText("18BR");
                                break;

                            case "19":
                                IWebElement nineteenBedrooms =
                                    waiting.Until(x => x.FindElement(By.Id("BedroomsDropDown")));
                                var act19 = new Actions(driver);
                                act19.MoveToElement(nineteenBedrooms)
                                    .Click(driver.FindElement(By.XPath("//*[@id='BedroomsDropDown']")))
                                    .Build()
                                    .Perform();
                                Thread.Sleep(1000);
                                var selectedBedroomNineteen = new SelectElement(nineteenBedrooms);
                                selectedBedroomNineteen.SelectByText("19BR");
                                break;

                            case "20":
                                IWebElement twentyBedrooms =
                                    waiting.Until(x => x.FindElement(By.Id("BedroomsDropDown")));
                                var act20 = new Actions(driver);
                                act20.MoveToElement(twentyBedrooms)
                                    .Click(driver.FindElement(By.XPath("//*[@id='BedroomsDropDown']")))
                                    .Build()
                                    .Perform();
                                Thread.Sleep(1000);
                                var selectedBedroomTwenty = new SelectElement(twentyBedrooms);
                                selectedBedroomTwenty.SelectByText("20BR");
                                break;

                            case "21":
                                IWebElement twentyOneBedrooms =
                                    waiting.Until(x => x.FindElement(By.Id("BedroomsDropDown")));
                                var act21 = new Actions(driver);
                                act21.MoveToElement(twentyOneBedrooms)
                                    .Click(driver.FindElement(By.XPath("//*[@id='BedroomsDropDown']")))
                                    .Build()
                                    .Perform();
                                Thread.Sleep(1000);
                                var selectedBedroomTwentyOne = new SelectElement(twentyOneBedrooms);
                                selectedBedroomTwentyOne.SelectByText("21BR");
                                break;

                            case "22":
                                IWebElement twentyTwoBedrooms =
                                    waiting.Until(x => x.FindElement(By.Id("BedroomsDropDown")));
                                var act22 = new Actions(driver);
                                act22.MoveToElement(twentyTwoBedrooms)
                                    .Click(driver.FindElement(By.XPath("//*[@id='BedroomsDropDown']")))
                                    .Build()
                                    .Perform();
                                Thread.Sleep(1000);
                                var selectedBedroomTwentyTwo = new SelectElement(twentyTwoBedrooms);
                                selectedBedroomTwentyTwo.SelectByText("22BR");
                                break;

                            case "23":
                                IWebElement twentyThreeBedrooms =
                                    waiting.Until(x => x.FindElement(By.Id("BedroomsDropDown")));
                                var act23 = new Actions(driver);
                                act23.MoveToElement(twentyThreeBedrooms)
                                    .Click(driver.FindElement(By.XPath("//*[@id='BedroomsDropDown']")))
                                    .Build()
                                    .Perform();
                                Thread.Sleep(1000);
                                var selectedBedroomTwentyThree = new SelectElement(twentyThreeBedrooms);
                                selectedBedroomTwentyThree.SelectByText("23BR");
                                break;

                            case "24":
                                IWebElement twentyFourBedrooms =
                                    waiting.Until(x => x.FindElement(By.Id("BedroomsDropDown")));
                                var act24 = new Actions(driver);
                                act24.MoveToElement(twentyFourBedrooms)
                                    .Click(driver.FindElement(By.XPath("//*[@id='BedroomsDropDown']")))
                                    .Build()
                                    .Perform();
                                Thread.Sleep(1000);
                                var selectedBedroomTwentyFour = new SelectElement(twentyFourBedrooms);
                                selectedBedroomTwentyFour.SelectByText("24BR");
                                break;

                            case "25":
                                IWebElement twentyFiveBedrooms =
                                    waiting.Until(x => x.FindElement(By.Id("BedroomsDropDown")));
                                var act25 = new Actions(driver);
                                act25.MoveToElement(twentyFiveBedrooms)
                                    .Click(driver.FindElement(By.XPath("//*[@id='BedroomsDropDown']")))
                                    .Build()
                                    .Perform();
                                Thread.Sleep(1000);
                                var selectedBedroomTwentyFive = new SelectElement(twentyFiveBedrooms);
                                selectedBedroomTwentyFive.SelectByText("25BR");
                                break;

                            case "26":
                                IWebElement twentySixBedrooms =
                                    waiting.Until(x => x.FindElement(By.Id("BedroomsDropDown")));
                                var act26 = new Actions(driver);
                                act26.MoveToElement(twentySixBedrooms)
                                    .Click(driver.FindElement(By.XPath("//*[@id='BedroomsDropDown']")))
                                    .Build()
                                    .Perform();
                                Thread.Sleep(1000);
                                var selectedBedroomTwentySix = new SelectElement(twentySixBedrooms);
                                selectedBedroomTwentySix.SelectByText("26BR");
                                break;

                            case "27":
                                IWebElement twentySevenBedrooms =
                                    waiting.Until(x => x.FindElement(By.Id("BedroomsDropDown")));
                                var act27 = new Actions(driver);
                                act27.MoveToElement(twentySevenBedrooms)
                                    .Click(driver.FindElement(By.XPath("//*[@id='BedroomsDropDown']")))
                                    .Build()
                                    .Perform();
                                Thread.Sleep(1000);
                                var selectedBedroomTwentySeven = new SelectElement(twentySevenBedrooms);
                                selectedBedroomTwentySeven.SelectByText("27BR");
                                break;

                            case "28":
                                IWebElement twentyEightBedrooms =
                                    waiting.Until(x => x.FindElement(By.Id("BedroomsDropDown")));
                                var act28 = new Actions(driver);
                                act28.MoveToElement(twentyEightBedrooms)
                                    .Click(driver.FindElement(By.XPath("//*[@id='BedroomsDropDown']")))
                                    .Build()
                                    .Perform();
                                Thread.Sleep(1000);
                                var selectedBedroomTwentyEight = new SelectElement(twentyEightBedrooms);
                                selectedBedroomTwentyEight.SelectByText("28BR");
                                break;

                            case "29":
                                IWebElement twentyNineBedrooms =
                                    waiting.Until(x => x.FindElement(By.Id("BedroomsDropDown")));
                                var act29 = new Actions(driver);
                                act29.MoveToElement(twentyNineBedrooms)
                                    .Click(driver.FindElement(By.XPath("//*[@id='BedroomsDropDown']")))
                                    .Build()
                                    .Perform();
                                Thread.Sleep(1000);
                                var selectedBedroomTwentyNine = new SelectElement(twentyNineBedrooms);
                                selectedBedroomTwentyNine.SelectByText("29BR");
                                break;

                            case "30":
                                IWebElement thirtyBedrooms =
                                    waiting.Until(x => x.FindElement(By.Id("BedroomsDropDown")));
                                var act30 = new Actions(driver);
                                act30.MoveToElement(thirtyBedrooms)
                                    .Click(driver.FindElement(By.XPath("//*[@id='BedroomsDropDown']")))
                                    .Build()
                                    .Perform();
                                Thread.Sleep(1000);
                                var selectedBedroomThirty = new SelectElement(thirtyBedrooms);
                                selectedBedroomThirty.SelectByText("30BR");
                                break;
                        }

                        // Unit Number

                        IWebElement unitNumber =
                            waiting.Until(x => x.FindElement(By.Id("ctl00_body_InputForm_UnitNumberTextField")));
                        unitNumber.SendKeys(uNumber);

                        // Unit Description

                        IWebElement iframe2 =
                            driver.FindElement(
                                By.XPath(
                                    "//*[@id='ctl00_body_InputForm_DescriptionLabel']/span[1]/div/div[2]/iframe"));
                        driver.SwitchTo().Frame(iframe2);

                        IWebElement unitDescription = waiting.Until(x => x.FindElement(By.TagName("body")));
                        unitDescription.SendKeys(uDescription);
                        driver.SwitchTo().DefaultContent();

                        // Bathrooms

                        IWebElement numberBathrooms =
                            waiting.Until(
                                x =>
                                    x.FindElement(
                                        By.Id("ctl00_body_InputForm_BathroomsControl_FullBathroomsDropDown")));
                        var selectedNumberBathrooms = new SelectElement(numberBathrooms);
                        selectedNumberBathrooms.SelectByText(numberBath);

                        // Sleeps

                        IWebElement sleeps =
                            waiting.Until(x => x.FindElement(By.Id("ctl00_body_InputForm_SleepsDropDown")));
                        var numberSleepSelected = new SelectElement(sleeps);
                        numberSleepSelected.SelectByText(numberGuess);

                        // Availability Settings

                        IWebElement inventoryMode =
                            waiting.Until(
                                x =>
                                    x.FindElement(
                                        By.Id("ctl00_body_InputForm_AvailabilitySettingsSection_InventoryModeField")));
                        var selectedInventoryMode = new SelectElement(inventoryMode);
                        selectedInventoryMode.SelectByText("Unit Specific");

                        IWebElement defaultAvailability =
                            waiting.Until(
                                x =>
                                    x.FindElement(
                                        By.Id(
                                            "ctl00_body_InputForm_AvailabilitySettingsSection_DefaultAvailabilityField")));
                        var selectedDefaultAvailability = new SelectElement(defaultAvailability);
                        selectedDefaultAvailability.SelectByText("Available");

                        IWebElement activelyManaged =
                            waiting.Until(
                                x =>
                                    x.FindElement(
                                        By.Id(
                                            "ctl00_body_InputForm_AvailabilitySettingsSection_IsActivelyManagedField")));
                        activelyManaged.Click();

                        // Saving Rental Unit

                        IWebElement finishButtonCreateRentalUnit =
                            driver.FindElement(By.Id("ctl00_body_InputForm_Finish"));
                        finishButtonCreateRentalUnit.Click();

                        // Begin Next Rental Unit

                        Thread.Sleep(1000);

                        if (driver.FindElement(By.Id("ctl00_PageTitleLabel")).Displayed)
                        {
                            IWebElement supplierLink =
                                waiting.Until(
                                    x =>
                                        x.FindElement(
                                            By.XPath("//*[@id='ctl00_PageContextText_ContextControl']/a[1]")));

                            Thread.Sleep(300);

                            var url =
                                waiting.Until(
                                    x =>
                                        x.FindElement(
                                            By.XPath(
                                                "//*[@id='ctl00_body_DashboardControl1_TaskSelection_ctl01_TaskGroup1_EditRentalUnitInfo']/a"))
                                            .GetAttribute("href"));

                            finalUrl = url.Substring((Math.Max(0, url.Length - 6)));

                            finalComplex = complex.Replace(",", " ");

                            Console.WriteLine(finalComplex + "," + building + "," + rooms + "," + uNumber + "," +
                                              uDescription + "," + numberBath + "," + numberGuess + "," + finalUrl);

                            supplierLink.Click();

                        }


                        Thread.Sleep(200);

                        if (
                            driver.FindElement(
                                By.XPath(
                                    "//*[@id='ctl00_body_DashboardControl1_TaskSelection_SupplierTasks1_TaskGroup1']/legend"))
                                .Displayed)
                        {
                            Thread.Sleep(200);

                            IWebElement createStandaloneRentalUnit2 =
                                waiting.Until(
                                    x =>
                                        x.FindElement(
                                            By.Id(
                                                "ctl00_body_DashboardControl1_TaskSelection_ctl01_RentalUnitTaskGroup_RentalUnitTasksControl_CreateRentalUnitSubTaskGroup_CreateARentalUnitTaskLink")));
                            createStandaloneRentalUnit2.Click();
                        }

                        // Console Write (Final Complex)

                        var rentalUnit = new RentalUnitInformation(finalComplex, building, rooms, uNumber,
                            uDescription, numberBath, numberGuess, finalUrl);
                        rentalUnitInfos.Add(rentalUnit);

                        if (i == rowCount - 1)
                        {
                            xlWorkBook.Close();
                            xlApp.Quit();
                            Marshal.ReleaseComObject(xlWorksheet);
                            Marshal.ReleaseComObject(xlWorkBook);
                            Marshal.ReleaseComObject(xlApp);

                            driver.Quit();

                            string csvPath = Environment.GetFolderPath(Environment.SpecialFolder.Desktop);
                            var newCsvFile = new FileStream(csvPath + suppName + ".txt", FileMode.Create, FileAccess.Write);
                            var fileToWrite = new StreamWriter(newCsvFile);

                            foreach (var eachrentalunit in rentalUnitInfos)
                            {
                                fileToWrite.WriteLine(eachrentalunit);
                            }

                            fileToWrite.Close();

                            Console.WriteLine("Finished creating the document for " + suppName + " Total Units " +
                                              (rowCount - 1));
                            Console.ReadLine();

                        }

                    } // end for

                } // end try
                catch (Exception exHandle)
                {

                    Console.Write("Exception " + exHandle.Message);
                    Console.ReadLine();
                }

                finally
                {
                    GC.Collect();
                }


            } // end if

            else
            {
                Console.WriteLine("Unable to find File");
                Console.ReadLine();
            }

            } // end method


    


           
    }
}
