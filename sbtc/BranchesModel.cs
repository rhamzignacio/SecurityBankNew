using System;
using System.Collections.Generic;
using System.Drawing;
using System.Linq;
using System.Text;

namespace sbtc
{
    public class BranchesModel
    {
        public string BRSTN { get; set; }
        public string Address1 { get; set; }
        public string Address2 { get; set; }
        public string Address3 { get; set; }
        public string Address4 { get; set; }
        public string Address5 { get; set; }
        public string Address6 { get; set; }
        public Int64 LastNo_PA { get; set; }
        public Int64 LastNo_CA { get; set; }
        public Int64 LastNo_MC { get; set; }
        public Int64 LastNo_Power_PA { get; set; }
        public Int64 LastNo_Power_CA { get; set; }
        public Int64 LastNo_GC { get; set; }
        public Int64 LastNo_CheckOne_PA { get; set; }
        public Int64 LastNo_CheckOne_CA { get; set; }
        public Nullable<DateTime> ModifiedDate { get; set; }
        public int IfChanges { get; set; } //1 = Changed, 0= No Change Made
    }

    public class OrderModel
    {
        public string CheckType { get; set; }
        public string BRSTN { get; set; }
        public string AccountNo { get; set; }

        private string _name;
        public string Name
        {
            get
            {
                if (_name == null)
                    return "";
                else
                    return _name.ToUpper();
            }
            set
            {
                _name = value;
            }
        }

        private string _name2;
        public string Name2
        {
            get
            {
                if (_name2 == null)
                    return "";
                else
                    return _name2.ToUpper();
            }
            set
            {
                _name2 = value;
            }
        }
        public string FormType { get; set; }
        public int OrderQuantity { get; set; }
        public string Batch { get; set; }

        private string _address1;
        public string Address1
        {
            get
            {
                if (_address1 == null)
                    return "";
                else
                    return _address1.ToUpper();
            }
            set
            {
                _address1 = value;
            }
        }

        private string _address2;
        public string Address2
        {
            get
            {
                if (_address2 == null)
                    return "";
                else
                    return _address2.ToUpper();
            }
            set
            {
                _address2 = value;
            }
        }

        private string _address3;
        public string Address3
        {
            get
            {
                if (_address3 == null)
                    return "";
                else
                    return _address3;
            }
            set
            {
                _address3 = value;
            }
        }

        private string _address4;
        public string Address4
        {
            get
            {
                if (_address4 == null)
                    return "";
                else
                    return _address4;
            }
            set
            {
                _address4 = value;
            }
        }

        private string _address5;
        public string Address5
        {
            get
            {
                if (_address5 == null)
                    return "";
                else
                    return _address5;
            }
            set
            {
                _address5 = value;
            }
        }

        private string _address6;
        public string Address6
        {
            get
            {
                if (_address6 == null)
                    return "";
                else
                    return _address6;
            }
            set
            {
                _address6 = value;
            }
        }
        public string ContCode { get; set; }
        public string CheckTypeName { get; set; }
        public Int64 StartingSerial { get; set; }
        public Int64 EndingSerial { get; set; }

        public Int64 ManualStart { get; set; }

        public string FileName { get; set; }
    }

    public class OrderSorted
    {
        public List<OrderModel> RegularPersonal { get; set; }
        public List<OrderModel> RegularCommercial { get; set; }
        public List<OrderModel> ManagersCheck { get; set; }
        public List<OrderModel> GiftCheck { get; set; }
        public List<OrderModel> PersonalPreEncoded { get; set; }
        public List<OrderModel> CommercialPreEncoded { get; set; }
        public List<OrderModel> CheckOnePersonal { get; set; }
        public List<OrderModel> CheckOneCommerical { get; set; }
        public List<OrderModel> CheckPowerPersonal { get; set; }
        public List<OrderModel> CheckPowerCommercial { get; set; }
        public List<OrderModel> CustomizedCheck { get; set; }
        public List<OrderModel> ManagersCheckCont { get; set; }
        public List<OrderModel> DigiBanker { get; set; }
        public List<OrderModel> Dividend { get; set; }
    }

    public class Locator
    { 
        public int PrimaryKey { get; set; }
        public string Location { get; set; }
    }
}
