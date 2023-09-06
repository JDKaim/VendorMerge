public class Consolidator
{
    public Consolidator()
    {

    }

    public VendorCollection Consolidate(VendorCollection master, VendorCollection competitor, Dictionary<string, string> renamers)
    {
        VendorCollection final = new VendorCollection("Final Sheet", renamers);
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
                                        if (product != "VHOSTPRO" && product != "PROSERV" && product != "PROWRK" && product != "PROSERV")
                                        {
                                            final.AddCustomerRecordQuantity(vendor.Vendor, customer.Customer, product, Math.Max(customer.GetQuantity(product), competingCustomer.GetQuantity(product)));
                                        }
                                        else
                                        {
                                            final.AddCustomerRecordQuantity(vendor.Vendor, customer.Customer, product, competingCustomer.GetQuantity(product));
                                        }
                                        competeContainsCustomer = true;
                                        break;
                                    }
                                }
                                if (!competeContainsCustomer)
                                {
                                    if (product != "VHOSTPRO" && product != "PROSERV" && product != "PROWRK" && product != "PROSERV")
                                    {
                                        final.AddCustomerRecordQuantity(vendor.Vendor, customer.Customer, product, customer.GetQuantity(product));
                                    }
                                    final.AddCustomerRecordQuantity(vendor.Vendor, customer.Customer, product, 0);
                                }
                            }
                            else
                            {
                                if (product != "VHOSTPRO" && product != "PROSERV" && product != "PROWRK" && product != "PROSERV")
                                {
                                    final.AddCustomerRecordQuantity(vendor.Vendor, customer.Customer, product, customer.GetQuantity(product));
                                }
                                final.AddCustomerRecordQuantity(vendor.Vendor, customer.Customer, product, 0);
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
                        if (product != "VHOSTPRO" && product != "PROSERV" && product != "PROWRK" && product != "PROSERV")
                        {
                            final.AddCustomerRecordQuantity(vendor.Vendor, customer.Customer, product, customer.GetQuantity(product));
                        }
                        final.AddCustomerRecordQuantity(vendor.Vendor, customer.Customer, product, 0);
                    }
                }
            }
        }
        return final;
    }
}