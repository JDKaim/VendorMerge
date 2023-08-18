public class Consolidator
{
    public Consolidator()
    {

    }

    public FinalCollection Consolidate(VendorCollection master, CompetingVendorCollection competitor)
    {
        FinalCollection final = new FinalCollection();
        foreach (VendorDataSet vendor in master.GetVendorDataSets())
        {
            bool inCompete = false;
            foreach (VendorDataSet competingVendor in competitor.GetVendorDataSets())
            {
                if (competingVendor.Vendor == vendor.Vendor)
                {
                    foreach (CustomerVendorRecord customer in vendor.GetCustomerVendorRecords())
                    {
                        foreach (string product in vendor.GetProducts())
                        {
                            if (competingVendor.GetProducts().Contains(product))
                            {
                                bool competeContainsCustomer = false;
                                foreach (CustomerVendorRecord competingCustomer in competingVendor.GetCustomerVendorRecords())
                                {
                                    if (competingCustomer.Customer == customer.Customer)
                                    {
                                        final.AddCustomerRecordQuantity(vendor.Vendor, customer.Customer, product, Math.Max(customer.GetQuantity(product), competingCustomer.GetQuantity(product)));
                                        competeContainsCustomer = true;
                                        break;
                                    }
                                }
                                if (!competeContainsCustomer)
                                {
                                    final.AddCustomerRecordQuantity(vendor.Vendor, customer.Customer, product, customer.GetQuantity(product));
                                }
                            }
                            else
                            {
                                final.AddCustomerRecordQuantity(vendor.Vendor, customer.Customer, product, customer.GetQuantity(product));
                            }
                        }
                    }
                    inCompete = true;
                    break;
                }
            }
            if (!inCompete)
            {
                foreach (CustomerVendorRecord customer in vendor.GetCustomerVendorRecords())
                {
                    foreach (string product in vendor.GetProducts())
                    {
                        final.AddCustomerRecordQuantity(vendor.Vendor, customer.Customer, product, customer.GetQuantity(product));
                    }
                }
            }
        }
        return final;
    }
}