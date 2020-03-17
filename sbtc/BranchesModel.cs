using System;
using System.Collections.Generic;
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
        public DateTime? ModifiedDate { get; set; }
        public int IfChanges { get; set; } //1=Changed, 0=No Change Made
    }

    public class OrderModel
    {
        public string CheckType { get; set; }
        public string BRSTN { get; set; }
        public string AccountNo { get; set; }
        public string Name { get; set; }
        public string Name2 { get; set; }
        public string FormType { get; set; }
        public int OrderQuantity { get; set; }
        public string Batch { get; set; }
        public string Address1 { get; set; }
        public string Address2 { get; set; }
        public string Address3 { get; set; }
        public string Address4 { get; set; }
        public string Address5 { get; set; }
        public string Address6 { get; set; }
        public string ContCode { get; set; }
        public string CheckTypeName { get; set; }
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
    }
}
