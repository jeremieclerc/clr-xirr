using Microsoft.VisualStudio.TestTools.UnitTesting;
using System;
using System.Data.SqlTypes;
using System.Data.SqlClient;

namespace XirrTest
{
    [TestClass]
    public class XirrTestSuite
    {
        [TestMethod]
        public void TestXIRRCalculation_WithPositiveCashFlows()
        {
            // Arrange
            var xirr = new XIRR();
            xirr.Init();

            // Define test data: initial investment and subsequent cash inflows
            var cashFlows = new (DateTime date, double value)[]
            {
                (new DateTime(2020, 1, 1), -10000.0), // Initial investment
                (new DateTime(2020, 4, 1), 2500.0),
                (new DateTime(2020, 7, 1), 2500.0),
                (new DateTime(2020, 10, 1), 2500.0),
                (new DateTime(2021, 1, 1), 2500.0)
            };

            double expectedXIRR = 0; // Expected XIRR value calculated externally

            // Act
            foreach (var cashFlow in cashFlows)
            {
                SqlDateTime date = new SqlDateTime(cashFlow.date);
                SqlDouble value = new SqlDouble(cashFlow.value);

                xirr.Accumulate(value, date);
            }

            SqlDouble result = xirr.Terminate();

            // Assert
            Assert.AreEqual(expectedXIRR, result.Value, 0.0001, "XIRR calculation did not match expected value.");
        }

        [TestMethod]
        public void TestXIRRCalculation_WithNegativeCashFlows()
        {
            // Arrange
            var xirr = new XIRR();
            xirr.Init();

            // Define test data: initial cash inflow and subsequent cash outflows
            var cashFlows = new (DateTime date, double value)[]
            {
                (new DateTime(2020, 1, 1), 10000.0), // Initial cash inflow
                (new DateTime(2020, 4, 1), -3000.0),
                (new DateTime(2020, 7, 1), -4000.0),
                (new DateTime(2020, 10, 1), -3500.0)
            };

            double expectedXIRR = 0.1004; // Expected XIRR value calculated externally

            // Act
            foreach (var cashFlow in cashFlows)
            {
                SqlDateTime date = new SqlDateTime(cashFlow.date);
                SqlDouble value = new SqlDouble(cashFlow.value);

                xirr.Accumulate(value, date);
            }

            SqlDouble result = xirr.Terminate();

            // Assert
            Assert.AreEqual(expectedXIRR, result.Value, 0.0001, "XIRR calculation did not match expected value.");
        }

        [TestMethod]
        public void TestXIRRCalculation_WithZeroAmounts()
        {
            // Arrange
            var xirr = new XIRR();
            xirr.Init();

            // Define test data including zero amounts
            var cashFlows = new (DateTime date, double value)[]
            {
                (new DateTime(2020, 1, 1), -10000.0),
                (new DateTime(2020, 4, 1), 0.0),     // Zero amount, should be ignored
                (new DateTime(2020, 7, 1), 5000.0),
                (new DateTime(2021, 1, 1), 6000.0)
            };

            double expectedXIRR = 0.1318;

            // Act
            foreach (var cashFlow in cashFlows)
            {
                SqlDateTime date = new SqlDateTime(cashFlow.date);
                SqlDouble value = new SqlDouble(cashFlow.value);

                xirr.Accumulate(value, date);
            }

            SqlDouble result = xirr.Terminate();

            // Assert
            Assert.AreEqual(expectedXIRR, result.Value, 0.0001, "XIRR calculation did not match expected value when zero amounts are present.");
        }

        [TestMethod]
        public void TestXIRRCalculation_WithDuplicateDates()
        {
            // Arrange
            var xirr = new XIRR();
            xirr.Init();

            // Define test data with duplicate dates
            var cashFlows = new (DateTime date, double value)[]
            {
                (new DateTime(2020, 1, 1), -5000.0),
                (new DateTime(2020, 1, 1), -5000.0), // Duplicate date, amounts should be summed
                (new DateTime(2020, 6, 1), 6000.0),
                (new DateTime(2020, 12, 1), 6000.0)
            };

            double expectedXIRR = 0.3190;

            // Act
            foreach (var cashFlow in cashFlows)
            {
                SqlDateTime date = new SqlDateTime(cashFlow.date);
                SqlDouble value = new SqlDouble(cashFlow.value);

                xirr.Accumulate(value, date);
            }

            SqlDouble result = xirr.Terminate();

            // Assert
            Assert.AreEqual(expectedXIRR, result.Value, 0.0001, "XIRR calculation did not match expected value with duplicate dates.");
        }

        [TestMethod]
        public void TestXIRRCalculation_NoPositiveCashFlows()
        {
            // Arrange
            var xirr = new XIRR();
            xirr.Init();

            // Define test data with only negative cash flows
            var cashFlows = new (DateTime date, double value)[]
            {
                (new DateTime(2020, 1, 1), -5000.0),
                (new DateTime(2020, 6, 1), -6000.0)
            };

            // Act
            foreach (var cashFlow in cashFlows)
            {
                SqlDateTime date = new SqlDateTime(cashFlow.date);
                SqlDouble value = new SqlDouble(cashFlow.value);

                xirr.Accumulate(value, date);
            }

            SqlDouble result = xirr.Terminate();

            // Assert
            Assert.IsTrue(result.IsNull, "Expected Null result when no positive cash flows are present.");
        }

        [TestMethod]
        public void TestXIRRCalculation_NoNegativeCashFlows()
        {
            // Arrange
            var xirr = new XIRR();
            xirr.Init();

            // Define test data with only positive cash flows
            var cashFlows = new (DateTime date, double value)[]
            {
                (new DateTime(2020, 1, 1), 5000.0),
                (new DateTime(2020, 6, 1), 6000.0)
            };

            // Act
            foreach (var cashFlow in cashFlows)
            {
                SqlDateTime date = new SqlDateTime(cashFlow.date);
                SqlDouble value = new SqlDouble(cashFlow.value);

                xirr.Accumulate(value, date);
            }

            SqlDouble result = xirr.Terminate();

            // Assert
            Assert.IsTrue(result.IsNull, "Expected Null result when no negative cash flows are present.");
        }
      

        [TestMethod]
        public void TestXIRRCalculation_HighPrecision()
        {
            // Arrange
            var xirr = new XIRR();
            xirr.Init();

            // Define test data with duplicate dates
            var cashFlows = new (DateTime date, double value)[]
            {
                (new DateTime(2016, 07, 28), -18630000.0000),
                (new DateTime(2016, 07, 28), -176594.0700),
                (new DateTime(2016, 08, 10), -11294.9100),
                (new DateTime(2016, 08, 10), -181541.2800),
                (new DateTime(2016, 08, 10), -295906.1200),
                (new DateTime(2016, 08, 10), -323300.0000),
                (new DateTime(2016, 08, 10), -299436.8000),
                (new DateTime(2016, 08, 10), -1344.0500),
                (new DateTime(2016, 08, 10), -15995.0900),
                (new DateTime(2016, 09, 15), -843752.0000),
                (new DateTime(2016, 09, 28), -3582.1800),
                (new DateTime(2016, 10, 03), -58690.3000),
                (new DateTime(2016, 10, 03), -19032.0000),
                (new DateTime(2016, 11, 23), -187101.0500),
                (new DateTime(2016, 11, 23), -32898.7200),
                (new DateTime(2016, 11, 25), -6767.0400),
                (new DateTime(2016, 12, 15), -19200.0000),
                (new DateTime(2017, 01, 17), -189100.0000),
                (new DateTime(2017, 02, 10), -64.3500),
                (new DateTime(2017, 02, 27), -2460.8100),
                (new DateTime(2017, 04, 12), -6580.6300),
                (new DateTime(2017, 06, 07), -2410.2000),
                (new DateTime(2017, 07, 28), -16200.0000),
                (new DateTime(2017, 08, 04), -8934.0000),
                (new DateTime(2017, 09, 08), -8448.9600),
                (new DateTime(2017, 09, 15), -1000000.0000),
                (new DateTime(2018, 01, 05), -821.8800),
                (new DateTime(2018, 01, 15), -821.8800),
                (new DateTime(2018, 10, 12), -500000.0000),
                (new DateTime(2019, 02, 22), -20852.4200),
                (new DateTime(2020, 05, 04), -8548.0000),
                (new DateTime(2020, 05, 19), -606.3000),
                (new DateTime(2020, 05, 19), -182104.0000),
                (new DateTime(2020, 06, 08), -16871.4000),
                (new DateTime(2020, 07, 31), -300000.0000),
                (new DateTime(2020, 09, 23), -147734.1000),
                (new DateTime(2020, 09, 23), -6783.5100),
                (new DateTime(2020, 09, 23), -111.1500),
                (new DateTime(2020, 10, 13), -6545.0000),
                (new DateTime(2020, 11, 24), -2300.0000),
                (new DateTime(2021, 05, 19), -1211.7800),
                (new DateTime(2021, 06, 17), -6871.2500),
                (new DateTime(2021, 07, 20), -26982.7300),
                (new DateTime(2021, 07, 23), -1426.0800),
                (new DateTime(2022, 05, 24), -6867.8400),
                (new DateTime(2022, 11, 24), -546859.0000),
                (new DateTime(2023, 03, 29), -7417.3600),
                (new DateTime(2023, 07, 05), -79300.0000),
                (new DateTime(2023, 07, 05), -3414.3000),
                (new DateTime(2023, 07, 28), -29303.2000),
                (new DateTime(2023, 08, 02), -91033.2271),
                (new DateTime(2023, 11, 30), -79300.0000),
                (new DateTime(2024, 03, 13), -15089.6700),
                (new DateTime(2024, 06, 05), -474828.0000),
                (new DateTime(2016, 07, 28), -50000000.0000),
                (new DateTime(2016, 07, 28), -19667037.0600),
                (new DateTime(2020, 04, 27), -444076625.8800),
                (new DateTime(2020, 05, 04), -3273858.8400),
                (new DateTime(2021, 03, 24), -179999505.5481),
                (new DateTime(2021, 03, 24), -493.9437),
                (new DateTime(2024, 05, 30), -243414176.0000),
                (new DateTime(2017, 07, 24), -10300000.0000),
                (new DateTime(2018, 09, 27), -9000000.0000),
                (new DateTime(2020, 04, 27), 15070634.9252),
                (new DateTime(2020, 04, 27), 1554364.3100),
                (new DateTime(2024, 06, 30), 1477999129.2025),
                (new DateTime(2024, 06, 30), 658.2281),
            };

            double expectedXIRR = 0.1220;
            ;

            // Act
            foreach (var cashFlow in cashFlows)
            {
                SqlDateTime date = new SqlDateTime(cashFlow.date);
                SqlDouble value = new SqlDouble(cashFlow.value);

                xirr.Accumulate(value, date);
            }

            SqlDouble result = xirr.Terminate();

            // Assert
            Assert.AreEqual(expectedXIRR, result.Value, 0.0001, "XIRR calculation did not match expected value with duplicate dates.");
        }

        [TestMethod]
        public void TestXIRRCalculation_HorizontalTangent()
        {
            // Arrange
            var xirr = new XIRR();
            xirr.Init();

            // Define test data with duplicate dates
            var cashFlows = new (DateTime date, double value)[]
            {
                (new DateTime(2002, 04, 15), -18231.0000),
                (new DateTime(2002, 06, 21), -45148.8200),
                (new DateTime(2002, 08, 14), -51001.6000),
                (new DateTime(2002, 09, 26), -12500.3900),
                (new DateTime(2002, 11, 13), 3958.4400),
                (new DateTime(2002, 11, 13), -3958.4431),
                (new DateTime(2002, 11, 13), 10375.3200),
                (new DateTime(2002, 12, 19), -19167.2800),
                (new DateTime(2003, 03, 27), -63503.2200),
                (new DateTime(2003, 05, 05), -17708.9000),
                (new DateTime(2003, 05, 30), -4166.7800),
                (new DateTime(2003, 06, 20), -11250.3500),
                (new DateTime(2003, 07, 10), 4166.7800),
                (new DateTime(2003, 07, 10), 11250.3516),
                (new DateTime(2003, 07, 10), -4166.7800),
                (new DateTime(2003, 08, 05), -466681.2300),
                (new DateTime(2003, 09, 10), 124126.3700),
                (new DateTime(2003, 09, 12), 10417.0000),
                (new DateTime(2003, 09, 12), -43751.3700),
                (new DateTime(2003, 09, 12), -10417.0000),
                (new DateTime(2003, 10, 13), -6250.2100),
                (new DateTime(2003, 10, 13), 6250.2100),
                (new DateTime(2003, 10, 13), 52501.6500),
                (new DateTime(2003, 11, 19), -125837.2600),
                (new DateTime(2003, 12, 19), -16667.1700),
                (new DateTime(2003, 12, 19), -8333.6000),
                (new DateTime(2003, 12, 19), 8333.6000),
                (new DateTime(2004, 01, 05), -37501.1600),
                (new DateTime(2004, 01, 22), 5833.5300),
                (new DateTime(2004, 01, 22), -125420.5800),
                (new DateTime(2004, 01, 22), -5833.5300),
                (new DateTime(2004, 02, 26), -8958.6200),
                (new DateTime(2004, 02, 26), 8958.6200),
                (new DateTime(2004, 02, 26), 89086.1000),
                (new DateTime(2004, 04, 14), -3.7400),
                (new DateTime(2004, 04, 19), -215472.0880),
                (new DateTime(2004, 04, 19), 646416.2642),
                (new DateTime(2001, 05, 31), -10000.0000),
                (new DateTime(2004, 04, 19), -1880.4120),
                (new DateTime(2004, 04, 19), 5641.2358),
                (new DateTime(2002, 01, 17), -39353.0000),
                (new DateTime(2002, 01, 17), -65.0000),
                (new DateTime(2002, 01, 17), 65.0000)
            };

            double expectedXIRR = -0.5252;
            ;

            // Act
            foreach (var cashFlow in cashFlows)
            {
                SqlDateTime date = new SqlDateTime(cashFlow.date);
                SqlDouble value = new SqlDouble(cashFlow.value);

                xirr.Accumulate(value, date);
            }

            SqlDouble result = xirr.Terminate();

            // Assert
            Assert.AreEqual(expectedXIRR, result.Value, 0.0001, "XIRR calculation did not match expected value with duplicate dates.");
        }

        [TestMethod]
        public void TestXIRRCalculation_Error()
        {
            // Arrange
            var xirr = new XIRR();
            xirr.Init();

            // Define test data with duplicate dates
            var cashFlows = new (DateTime date, double value)[]
            {
                (new DateTime(2022, 10, 06), -341832.2850),
                (new DateTime(2022, 10, 06), -0.7587),
                (new DateTime(2022, 07, 29), -2392825.9950),
                (new DateTime(2022, 09, 30), -8670258774.1226),
                (new DateTime(2022, 09, 30), -87578371.4558),
                (new DateTime(2022, 09, 30), 1466594881.6795),
                (new DateTime(2022, 09, 30), -14309381461.7351),
                (new DateTime(2023, 01, 01), -36116320.8366),
                (new DateTime(2023, 02, 09), -12037832771.7675),
                (new DateTime(2023, 06, 28), -341832285.0000),
                (new DateTime(2023, 11, 21), -581114884.5000),
                (new DateTime(2023, 11, 21), -305092181.7731),
                (new DateTime(2023, 11, 21), -136732914.0000),
                (new DateTime(2024, 04, 02), -136732914.0000),
                (new DateTime(2024, 06, 30), 116231300.8627),
                (new DateTime(2024, 06, 30), 11506898785.4025),
                (new DateTime(2024, 06, 30), 5623100761.7583),
                (new DateTime(2024, 06, 30), 12776530401.7733),
                (new DateTime(2024, 06, 30), 15281031521.6065)
            };

            double expectedXIRR = 0.1729;


            // Act
            foreach (var cashFlow in cashFlows)
            {
                SqlDateTime date = new SqlDateTime(cashFlow.date);
                SqlDouble value = new SqlDouble(cashFlow.value);

                xirr.Accumulate(value, date);
            }

            SqlDouble result = xirr.Terminate();

            // Assert
            Assert.AreEqual(expectedXIRR, result.Value, 0.0001, "XIRR calculation did not match expected value with duplicate dates.");
        }

        [TestMethod]
        public void TestXIRRCalculation_Error2()
        {
            // Arrange
            var xirr = new XIRR();
            xirr.Init();

            // Define test data with duplicate dates
            var cashFlows = new (DateTime date, double value)[]
            {
                (new DateTime(2020, 04, 24), -368263.3300),
                (new DateTime(2020, 04, 29), 24642.7400),
                (new DateTime(2020, 04, 29), 2465.7500),
                (new DateTime(2021, 03, 23), -106961.2000),
                (new DateTime(2021, 07, 27), 1167392.5211),
                (new DateTime(2021, 07, 27), 3339199.6319),
                (new DateTime(2021, 09, 14), 2226131.9200),
                (new DateTime(2021, 09, 14), 786238.5400),
                (new DateTime(2020, 04, 24), -368263.3300),
                (new DateTime(2020, 04, 29), 24642.7400),
                (new DateTime(2020, 04, 29), 2465.7500),
                (new DateTime(2021, 03, 23), -106961.2000),
                (new DateTime(2021, 07, 27), 1167392.5211),
                (new DateTime(2021, 07, 27), 3339199.6319),
                (new DateTime(2021, 09, 14), 2226131.9200),
                (new DateTime(2021, 09, 14), 786238.5400),
                (new DateTime(2020, 04, 24), -368263.3300),
                (new DateTime(2020, 04, 29), 24642.7400),
                (new DateTime(2020, 04, 29), 2465.7600),
                (new DateTime(2021, 03, 23), -106961.2000),
                (new DateTime(2021, 07, 27), 1167392.5211),
                (new DateTime(2021, 07, 27), 3339199.6319),
                (new DateTime(2021, 09, 14), 2226131.9200),
                (new DateTime(2021, 09, 14), 786238.5400),
                (new DateTime(2022, 01, 05), -24919.8500),
                (new DateTime(2022, 02, 25), -27027.0000),
                (new DateTime(2020, 04, 24), -444640177.1200),
                (new DateTime(2020, 04, 29), 29753572.6000),
                (new DateTime(2021, 03, 23), -105891588.0000),
                (new DateTime(2020, 04, 24), -368263.3300),
                (new DateTime(2020, 04, 29), 24642.7400),
                (new DateTime(2020, 04, 29), 2465.7600),
                (new DateTime(2021, 03, 23), -106961.2000),
                (new DateTime(2021, 07, 27), 1167392.5211),
                (new DateTime(2021, 07, 27), 3339199.6319),
                (new DateTime(2021, 09, 14), 2226131.9200),
                (new DateTime(2021, 09, 14), 786238.5400),
                (new DateTime(2020, 04, 24), -368263.3300),
                (new DateTime(2020, 04, 29), 24642.7400),
                (new DateTime(2020, 04, 29), 2465.7500),
                (new DateTime(2021, 07, 27), 1167392.5211),
                (new DateTime(2021, 07, 27), 3339199.6319),
                (new DateTime(2021, 09, 14), 2226131.9200),
                (new DateTime(2021, 09, 14), 786238.5400),
                (new DateTime(2020, 04, 24), -368263.3300),
                (new DateTime(2020, 04, 29), 24642.7400),
                (new DateTime(2020, 04, 29), 2465.7600),
                (new DateTime(2021, 03, 23), -106961.2000),
                (new DateTime(2021, 07, 21), -59400.0000),
                (new DateTime(2021, 07, 27), 1167392.5211),
                (new DateTime(2021, 07, 27), 3339199.6319),
                (new DateTime(2021, 09, 14), 2226131.9200),
                (new DateTime(2021, 09, 14), 786238.5400),
                (new DateTime(2021, 09, 28), -221076.8700),
                (new DateTime(2021, 09, 28), -821689.5500),
                (new DateTime(2021, 11, 04), -1631.6300),
                (new DateTime(2021, 11, 04), -1688.8300),
                (new DateTime(2021, 11, 04), -1640.9800),
                (new DateTime(2021, 11, 04), -1631.6300),
                (new DateTime(2021, 12, 17), -5954.6400),
                (new DateTime(2023, 04, 28), -39677.8200),
                (new DateTime(2020, 04, 24), -368263.3300),
                (new DateTime(2020, 04, 29), 24642.7400),
                (new DateTime(2020, 04, 29), 2465.7500),
                (new DateTime(2021, 03, 23), -106961.2000),
                (new DateTime(2021, 07, 27), 1167392.5211),
                (new DateTime(2021, 07, 27), 3339199.6319),
                (new DateTime(2021, 09, 14), 2226131.9200),
                (new DateTime(2021, 09, 14), 786238.5400),
                (new DateTime(2020, 04, 24), -368263.3300),
                (new DateTime(2020, 04, 29), 24642.7400),
                (new DateTime(2020, 04, 29), 2465.7500),
                (new DateTime(2021, 03, 23), -106961.2000),
                (new DateTime(2021, 03, 23), -106961.2000),
                (new DateTime(2021, 07, 27), 1167392.5211),
                (new DateTime(2021, 07, 27), 3339199.6319),
                (new DateTime(2021, 09, 14), 2226131.9200),
                (new DateTime(2021, 09, 14), 786238.5400),
                (new DateTime(2020, 04, 24), -368263.3300),
                (new DateTime(2020, 04, 29), 24642.7400),
                (new DateTime(2020, 04, 29), 2465.7500),
                (new DateTime(2021, 03, 23), -106961.2000),
                (new DateTime(2021, 07, 27), 1167392.5211),
                (new DateTime(2021, 07, 27), 3339199.6319),
                (new DateTime(2021, 09, 14), 2226131.9200),
                (new DateTime(2021, 09, 14), 786238.5400),
                (new DateTime(2020, 04, 24), -368263.3300),
                (new DateTime(2020, 04, 29), 24642.7400),
                (new DateTime(2020, 04, 29), 2465.7500),
                (new DateTime(2021, 03, 23), -106961.2000),
                (new DateTime(2021, 07, 27), 1167392.5211),
                (new DateTime(2021, 07, 27), 3339199.6319),
                (new DateTime(2021, 09, 14), 2226131.9200),
                (new DateTime(2021, 09, 14), 786238.5400)
            };

            double expectedXIRR = -0.8773;

            var aggregatedCashFlows = cashFlows
               .GroupBy(c => c.date)
               .Select(g => (date: g.Key, value: g.Sum(c => c.value)))
               .ToList();


            // Act
            foreach (var cashFlow in aggregatedCashFlows)
            {
                SqlDateTime date = new SqlDateTime(cashFlow.date);
                SqlDouble value = new SqlDouble(cashFlow.value);

                xirr.Accumulate(value, date);
            }

            SqlDouble result = xirr.Terminate();

            // Assert
            Assert.AreEqual(expectedXIRR, result.Value, 0.0001, "XIRR calculation did not match expected value with duplicate dates.");
        }

        [TestMethod]
        public void TestXIRRCalculation_Error3()
        {
            // Arrange
            var xirr = new XIRR();
            xirr.Init();

            // Define test data with duplicate dates
            var cashFlows = new (DateTime date, double value)[]
            {
                (new DateTime(2004, 08, 27), -205719.0000),
                (new DateTime(2004, 10, 19), -109297.1400),
                (new DateTime(2004, 12, 15), -47389.0000),
                (new DateTime(2004, 12, 31), -12365.7000),
                (new DateTime(2005, 03, 09), -116182.0600),
                (new DateTime(2005, 04, 22), 45357.0600),
                (new DateTime(2005, 05, 20), -67998.7400),
                (new DateTime(2005, 08, 30), -58739.5300),
                (new DateTime(2005, 10, 03), -93426.5500),
                (new DateTime(2005, 10, 26), -264985.5900),
                (new DateTime(2005, 11, 07), 266572.3300),
                (new DateTime(2005, 12, 02), -146190.3600),
                (new DateTime(2005, 12, 22), 564728.7300),
                (new DateTime(2006, 06, 02), 152887.4600),
                (new DateTime(2006, 06, 26), -167688.0300),
                (new DateTime(2006, 07, 25), 286277.8400),
                (new DateTime(2006, 08, 22), 275706.2900),
                (new DateTime(2006, 09, 27), -147567.0000),
                (new DateTime(2006, 10, 10), 220519.2900),
                (new DateTime(2006, 12, 01), -153584.3300),
                (new DateTime(2007, 01, 31), 499303.5400),
                (new DateTime(2007, 03, 30), 148875.9500),
                (new DateTime(2007, 04, 23), 372080.9077),
                (new DateTime(2007, 05, 31), 216819.4646),
                (new DateTime(2007, 12, 06), -42719.3000),
                (new DateTime(2008, 07, 24), -156202.0000),
                (new DateTime(2008, 09, 12), 223424.0000),
                (new DateTime(2008, 09, 24), 486530.3700),
                (new DateTime(2008, 11, 14), 47722.5700),
                (new DateTime(2009, 07, 17), 61159.6200),
                (new DateTime(2010, 05, 18), 21165.9800),
                (new DateTime(2010, 10, 04), 606084.7200),
                (new DateTime(2010, 10, 29), 217160.6600),
                (new DateTime(2011, 01, 06), 212419.7600),
                (new DateTime(2011, 02, 25), 34482.3100),
                (new DateTime(2011, 04, 04), 183017.5700),
                (new DateTime(2011, 06, 22), 712042.6500),
                (new DateTime(2011, 07, 14), 110597.1200),
                (new DateTime(2011, 11, 17), 27936.9000),
                (new DateTime(2011, 12, 07), 48357.5700),
                (new DateTime(2012, 05, 25), 348400.8900),
                (new DateTime(2012, 08, 03), 41471.5900),
                (new DateTime(2012, 09, 28), 8359.4600),
                (new DateTime(2012, 12, 27), 98824.5400),
                (new DateTime(2013, 06, 18), 43677.4800),
                (new DateTime(2013, 08, 21), 44643.6200),
                (new DateTime(2014, 05, 07), 191850.8500),
                (new DateTime(2014, 11, 26), 606182.1400),
                (new DateTime(2015, 10, 20), 321472.5200),
                (new DateTime(2017, 06, 07), 101637.5400),
                (new DateTime(2018, 01, 04), -4011706.1600)
            };

            double expectedXIRR = 1.077;

            var aggregatedCashFlows = cashFlows
               .GroupBy(c => c.date)
               .Select(g => (date: g.Key, value: g.Sum(c => c.value)))
               .ToList();


            // Act
            foreach (var cashFlow in aggregatedCashFlows)
            {
                SqlDateTime date = new SqlDateTime(cashFlow.date);
                SqlDouble value = new SqlDouble(cashFlow.value);

                xirr.Accumulate(value, date);
            }

            SqlDouble result = xirr.Terminate();

            // Assert
            Assert.AreEqual(expectedXIRR, result.Value, 0.0001, "XIRR calculation did not match expected value with duplicate dates.");
        }

        [TestMethod]
        public void TestXIRRCalculation_UnderMinusOne()
        {
            // Arrange
            var xirr = new XIRR();
            xirr.Init();

            // Define test data with duplicate dates
            var cashFlows = new (DateTime date, double value)[]
            {
               (new DateTime(2012, 10, 17), -790672.0000),
                (new DateTime(2012, 12, 21), 629491.0000),
                (new DateTime(2014, 05, 12), 161181.0000),
                (new DateTime(2012, 10, 10), -250.0000),
                (new DateTime(2012, 10, 17), -1854664.0000),
                (new DateTime(2012, 10, 17), -250.0000),
                (new DateTime(2012, 10, 31), 9183755.0000),
                (new DateTime(2012, 12, 12), -851.7200),
                (new DateTime(2012, 12, 21), 1596047.0000),
                (new DateTime(2013, 01, 21), -9432672.1500),
                (new DateTime(2013, 01, 21), 9432672.1500),
                (new DateTime(2013, 08, 02), -55000.0000),
                (new DateTime(2014, 03, 27), -6292.2400),
                (new DateTime(2014, 03, 27), -808.3300),
                (new DateTime(2014, 07, 29), 294564.7600),
                (new DateTime(2014, 08, 04), -120.0000),
                (new DateTime(2014, 10, 22), -27508.0000),
                (new DateTime(2014, 11, 25), -88900.5000),
                (new DateTime(2017, 09, 05), -354.0100),
                (new DateTime(2012, 10, 17), -99.0000),
                (new DateTime(2012, 10, 17), -96898695.0000),
                (new DateTime(2012, 10, 17), -891.0000),
                (new DateTime(2012, 10, 17), -10.0000),
                (new DateTime(2012, 10, 17), -1854664.0000),
                (new DateTime(2012, 12, 21), 1596047.0000),
                (new DateTime(2014, 07, 18), -600000.0000),
                (new DateTime(2014, 07, 29), 294564.7700),
                (new DateTime(2012, 10, 17), -790672.0000),
                (new DateTime(2012, 12, 21), 629491.0000),
                (new DateTime(2014, 05, 12), 161181.0000),
                (new DateTime(2012, 10, 10), -250.0000),
                (new DateTime(2012, 10, 17), -1854664.0000),
                (new DateTime(2012, 10, 17), -250.0000),
                (new DateTime(2012, 10, 31), 9183755.0000),
                (new DateTime(2012, 12, 12), -851.7200),
                (new DateTime(2012, 12, 21), 1596047.0000),
                (new DateTime(2013, 01, 21), -9432672.1500),
                (new DateTime(2013, 01, 21), 9432672.1500),
                (new DateTime(2013, 08, 02), -55000.0000),
                (new DateTime(2014, 03, 27), -6292.2400),
                (new DateTime(2014, 03, 27), -808.3300),
                (new DateTime(2014, 07, 29), 294564.7600),
                (new DateTime(2014, 08, 04), -120.0000),
                (new DateTime(2014, 10, 22), -27508.0000),
                (new DateTime(2014, 11, 25), -88900.5000),
                (new DateTime(2017, 09, 05), -354.0100),
                (new DateTime(2012, 10, 17), -99.0000),
                (new DateTime(2012, 10, 17), -96898695.0000),
                (new DateTime(2012, 10, 17), -891.0000),
                (new DateTime(2012, 10, 17), -10.0000),
                (new DateTime(2012, 10, 17), -1854664.0000),
                (new DateTime(2012, 12, 21), 1596047.0000),
                (new DateTime(2014, 07, 18), -600000.0000),
                (new DateTime(2014, 07, 29), 294564.7700)

            };

            var aggregatedCashFlows = cashFlows
               .GroupBy(c => c.date)
               .Select(g => (date: g.Key, value: g.Sum(c => c.value)))
               .ToList();

            // Act
            foreach (var cashFlow in aggregatedCashFlows)
            {
                SqlDateTime date = new SqlDateTime(cashFlow.date);
                SqlDouble value = new SqlDouble(cashFlow.value);

                xirr.Accumulate(value, date);
            }

            SqlDouble result = xirr.Terminate();

            // Assert
            Assert.IsTrue(result.IsNull, "Expected Null result when the calculated IRR is negative");
        }
    }
}
