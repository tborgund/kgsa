//using System;
//using System.Collections.Generic;
//using System.Linq;
//using System.Text;

//namespace KGSA
//{
//    public class KgsaBudget
//    {

//        private void Test()
//        {
//            var budget = new BudgetAvdeling();
//            budget.avdeling = KgsaAvdeling.Data;

//            budget.AddSelger("TRONDBO");
//        }
//    }


//    public class BudgetAvdeling
//    {
//        public KgsaAvdeling avdeling;
//        public int id;

//        public List<BudgetSelger> selgere;

//        public void AddSelger(string sk)
//        {
//            var item = new BudgetSelger();
//            item.selgerkode = sk;
//            item.timer = 0;
//            item.dager = 0;
//            selgere.Add(item);
//        }
//    }


//    public class BudgetSelger
//    {
//        public string selgerkode;
//        public int timer;
//        public int dager;
//        public decimal multiplikator;
//        public List<BudgetItem> items;

//        public void AddItem(BudgetType type)
//        {

//        }
//    }

//    public class BudgetItem
//    {
//        public BudgetType type;

//        public decimal inntjen;
//        public decimal omset;
//        public decimal som;
//        public decimal sob;

//        public BudgetTargetValueType target;

//        public decimal target;
//        public decimal diff;
//    }

//    public enum BudgetType { Overall, TA, Strom, Finans, Rtgsa, Acc }

//    public enum BudgetTargetValueType { SoM, SoB, Inntjening, Omsetning, Antall, Hitrate }

//    public enum KgsaAvdeling { MDA, AudioVideo, SDA, Tele, Data, Kasse, Aftersales }
//}
