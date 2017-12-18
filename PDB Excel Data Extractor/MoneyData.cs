using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.Linq;
using System.Reflection.Emit;
using System.Text;
using System.Threading.Tasks;

namespace PDB_Excel_Data_Extractor
{
    public class MoneyData
    {
        public  int Ammout(string givenMoney, string typeOfCard)
        {
            int ammountToLoad = 0;
            if (typeOfCard.ToLower().Equals("silver card") || typeOfCard.ToLower().Equals("golden card") ||
                typeOfCard.ToLower().Equals("platinum card"))
            {
                ammountToLoad = AmmountBaseOnTypeOfCard(typeOfCard, givenMoney);
            }
            else
            {
                ammountToLoad = NoSpecialCards(givenMoney);
            }
            return ammountToLoad;
        }
        private int NoSpecialCards(string givenMoney)
        {
            int amountOfMoney = 0;
            switch (givenMoney)
            {
                case "50":
                    amountOfMoney = 72;
                    break;
                case "60":
                    amountOfMoney = 72;
                    break;
                case "110":
                    amountOfMoney = 144;
                    break;
                case "80":
                case "155":
                    amountOfMoney = 216;
                    break;
                case "100":
                case "195":
                    amountOfMoney = 288;
                    break;
                case "125":
                case "240":
                    amountOfMoney = 360;
                    break;
                case "180":
                case "350":
                    amountOfMoney = 700;
                    break;
            }
            return amountOfMoney;
        }
        private  int AmmountBaseOnTypeOfCard(string type, string givenMoney)
        {
            int amountOfMoney = 0;
            switch (type)
            {
                case "Silver card":
                    switch (givenMoney)
                    {
                        case "110":
                            amountOfMoney = 162;
                            break;
                        case "80":
                        case "155":
                            amountOfMoney = 252;
                            break;
                        case "100":
                        case "195":
                            amountOfMoney = 324;
                            break;
                    }
                    break;
                case "Golden card":
                    switch (givenMoney)
                    {
                        case "110":
                            amountOfMoney = 180;
                            break;
                        case "80":
                        case "155":
                            amountOfMoney = 288;
                            break;
                        case "100":
                        case "195":
                            amountOfMoney = 360;
                            break;
                    }
                    break;
                case "Platinum card":
                    switch (givenMoney)
                    {
                        case "110":
                            amountOfMoney = 198;
                            break;
                        case "80":
                        case "155":
                            amountOfMoney = 324;
                            break;
                        case "100":
                        case "195":
                            amountOfMoney = 396;
                            break;
                    }
                    break;
            }
            return amountOfMoney;
        }
        //Check for the card's expiration period based on the ammount of money
        public string CardPeriodExpiration(int year, int month, int day, int givenMoney)
        {
            string date = "";
            var dat = new DateTime(year, month, day);
            if (givenMoney > 110)
            {
                date = dat.AddMonths(2).ToString("dd.MM.yyyy");
            }
            else
            {
                date = dat.AddMonths(1).ToString("dd.MM.yyyy");
            }
            return date;
        }
        //Checks if there is Deffered payment
        public string deferredPayment(string ammountOfMoney)
        {
            var typeOfPayment = "";
            switch (ammountOfMoney)
            {
                case "80":
                    typeOfPayment = "50%";
                    break;
                case "100":
                    typeOfPayment = "50%";
                    break;
                case "125":
                    typeOfPayment = "50%";
                    break;
                case "180":
                    typeOfPayment = "50%";
                    break;
                default:
                    typeOfPayment = "100%";
                    break;
            }
            return typeOfPayment;
        }

        public int RemainingSum()
        {
            int remainingSum = 0;
            return remainingSum;
        }

        //private List<int> GetColumIndex()
        //{
        //}



    }
}
