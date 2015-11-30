using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace HelloWPF
{
    public class RentalUnitInformation
    {

        public RentalUnitInformation(string complexName, string myBuilding, string unitRooms, string unitNum, string unitDescription, string unitNumberOfbath, string unitnumberOfPeople, string rentalUnitId)
        {

            ComplexName = complexName;
            Building = myBuilding;
            Rooms = unitRooms;
            UnitNumber = unitNum;
            UDescription = unitDescription;
            NumberOfBath = unitNumberOfbath;
            NumberOfPeople = unitnumberOfPeople;
            RentalUnitId = rentalUnitId;

        }

        public string ComplexName { get; set; }
        public string Description { get; set; }
        public string Building { get; set; }
        public string Rooms { get; set; }
        public string UnitNumber { get; set; }
        public string UDescription { get; set; }
        public string NumberOfBath { get; set; }
        public string NumberOfPeople { get; set; }
        public string RentalUnitId { get; set; }

        public override string ToString()
        {
            return ComplexName + "," + Building + "," + Rooms + "," + UnitNumber + "," + UDescription + "," + NumberOfBath + "," + NumberOfPeople + "," + RentalUnitId;
        }
    }
}
