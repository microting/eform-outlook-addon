using NUnit.Framework;
using OutlookSql;

namespace MicrotingOutlook.Integration.Tests
{
    [TestFixture]
    public class AppointmentTests : DbTestFixture
    {
        [Test]
        public void CanPeepTwice()
        {
            var appointment = new appointments();
            DbContext.appointments.Add(appointment);

            DbContext.SaveChanges();
        }
    }
}
