using Microsoft.VisualStudio.TestTools.UnitTesting;
using System;
using System.Data.SqlTypes;

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
            var cashFlows = new (DateTime Date, double Amount)[]
            {
                (new DateTime(2020, 1, 1), -10000.0), // Initial investment
                (new DateTime(2020, 4, 1), 2500.0),
                (new DateTime(2020, 7, 1), 2500.0),
                (new DateTime(2020, 10, 1), 2500.0),
                (new DateTime(2021, 1, 1), 2500.0)
            };

            double initialGuess = 0; // Initial approximation for IRR
            double expectedXIRR = 0; // Expected XIRR value calculated externally

            // Act
            foreach (var cashFlow in cashFlows)
            {
                SqlDateTime date = new SqlDateTime(cashFlow.Date);
                SqlDouble amount = new SqlDouble(cashFlow.Amount);
                SqlDouble irrApprox = new SqlDouble(initialGuess);

                xirr.Accumulate(date, amount, irrApprox);
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
            var cashFlows = new (DateTime Date, double Amount)[]
            {
                (new DateTime(2020, 1, 1), 10000.0), // Initial cash inflow
                (new DateTime(2020, 4, 1), -3000.0),
                (new DateTime(2020, 7, 1), -4000.0),
                (new DateTime(2020, 10, 1), -3500.0)
            };

            double initialGuess = 0.1; // Initial approximation for IRR
            double expectedXIRR = 0.0100; // Expected XIRR value calculated externally

            // Act
            foreach (var cashFlow in cashFlows)
            {
                SqlDateTime date = new SqlDateTime(cashFlow.Date);
                SqlDouble amount = new SqlDouble(cashFlow.Amount);
                SqlDouble irrApprox = new SqlDouble(initialGuess);

                xirr.Accumulate(date, amount, irrApprox);
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
            var cashFlows = new (DateTime Date, double Amount)[]
            {
                (new DateTime(2020, 1, 1), -10000.0),
                (new DateTime(2020, 4, 1), 0.0),     // Zero amount, should be ignored
                (new DateTime(2020, 7, 1), 5000.0),
                (new DateTime(2021, 1, 1), 6000.0)
            };

            double initialGuess = 0.1;
            double expectedXIRR = 0.1318;

            // Act
            foreach (var cashFlow in cashFlows)
            {
                SqlDateTime date = new SqlDateTime(cashFlow.Date);
                SqlDouble amount = new SqlDouble(cashFlow.Amount);
                SqlDouble irrApprox = new SqlDouble(initialGuess);

                xirr.Accumulate(date, amount, irrApprox);
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
            var cashFlows = new (DateTime Date, double Amount)[]
            {
                (new DateTime(2020, 1, 1), -5000.0),
                (new DateTime(2020, 1, 1), -5000.0), // Duplicate date, amounts should be summed
                (new DateTime(2020, 6, 1), 6000.0),
                (new DateTime(2020, 12, 1), 6000.0)
            };

            double initialGuess = 0.1;
            double expectedXIRR = 0.3190;
;

            // Act
            foreach (var cashFlow in cashFlows)
            {
                SqlDateTime date = new SqlDateTime(cashFlow.Date);
                SqlDouble amount = new SqlDouble(cashFlow.Amount);
                SqlDouble irrApprox = new SqlDouble(initialGuess);

                xirr.Accumulate(date, amount, irrApprox);
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
            var cashFlows = new (DateTime Date, double Amount)[]
            {
                (new DateTime(2020, 1, 1), -5000.0),
                (new DateTime(2020, 6, 1), -6000.0)
            };

            double initialGuess = 0.1;

            // Act
            foreach (var cashFlow in cashFlows)
            {
                SqlDateTime date = new SqlDateTime(cashFlow.Date);
                SqlDouble amount = new SqlDouble(cashFlow.Amount);
                SqlDouble irrApprox = new SqlDouble(initialGuess);

                xirr.Accumulate(date, amount, irrApprox);
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
            var cashFlows = new (DateTime Date, double Amount)[]
            {
                (new DateTime(2020, 1, 1), 5000.0),
                (new DateTime(2020, 6, 1), 6000.0)
            };

            double initialGuess = 0.1;

            // Act
            foreach (var cashFlow in cashFlows)
            {
                SqlDateTime date = new SqlDateTime(cashFlow.Date);
                SqlDouble amount = new SqlDouble(cashFlow.Amount);
                SqlDouble irrApprox = new SqlDouble(initialGuess);

                xirr.Accumulate(date, amount, irrApprox);
            }

            SqlDouble result = xirr.Terminate();

            // Assert
            Assert.IsTrue(result.IsNull, "Expected Null result when no negative cash flows are present.");
        }
    }
}
