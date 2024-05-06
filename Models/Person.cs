using Asp_MVC1.Models;

namespace Asp_MVC1.Models
{
    public class Person
    {
        public string FirstName { get; set; }
        public string LastName { get; set; }
        public string Gender { get; set; }
        public DateTime DateOfBirth { get; set; }
        public string PhoneNumber { get; set; }
        public string BirthPlace { get; set; }
        public bool IsGraduated { get; set; }
        public string FullName => LastName + ' ' + FirstName;
    }
}
