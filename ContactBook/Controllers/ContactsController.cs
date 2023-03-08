using System;
using System.Collections.Generic;
using System.Linq;
using System.Threading.Tasks;
using Microsoft.AspNetCore.Mvc;
using Microsoft.AspNetCore.Mvc.Rendering;
using Microsoft.EntityFrameworkCore;
using ContactBook.Data;
using ContactBook.Models;
using Microsoft.AspNetCore.Authorization;
using System.Text;
using ClosedXML.Excel;

namespace ContactBook.Controllers
{
    public class ContactsController : Controller
    {
        private readonly ApplicationDbContext _context;

        public ContactsController(ApplicationDbContext context)
        {
            _context = context;
        }

        [Authorize(Roles = "Admin, User")]
        // GET: Contacts
        public async Task<IActionResult> Index()
        {
              return View(await _context.Contacts.ToListAsync());
        }

        // Global variable
        private static List<Contacts> _contacts;

        //Search functionality
        [Authorize(Roles = "Admin, User")]
        [HttpGet]
        public async Task<IActionResult> Index(string ContactSearch)
        {
            ViewData["GetContactDetails"] = ContactSearch;
            var contactQuery = from x in _context.Contacts select x;

            if (!String.IsNullOrEmpty(ContactSearch))
            {
                contactQuery = contactQuery.Where(x => x.Name.Contains(ContactSearch) || x.Lastname.Contains(ContactSearch) ||
                                                  x.Address.Contains(ContactSearch) || x.Email.Contains(ContactSearch) || x.Phone.Contains(ContactSearch));
            }

            _contacts = await contactQuery.AsNoTracking().ToListAsync();
            return View(_contacts);
        }


        //Download to csv
        public IActionResult DownloadSearchResults()
        {

            if (_contacts != null)
            {

                var csv = new StringBuilder();
                csv.AppendLine("Id, First Name, Last Name, Phone Number, Email, Address");
                foreach (var contact in _contacts)
                {
                    csv.AppendLine($"{contact.ID},{contact.Name},{contact.Lastname},{contact.Phone},{contact.Email},{contact.Address}");
                }

                var csvBytes = Encoding.UTF8.GetBytes(csv.ToString());
                var csvFilename = "search_results.csv";
                var csvMimeType = "text/csv";
                return File(csvBytes, csvMimeType, csvFilename);
            }
            else
            {
                return RedirectToAction("Index");
            }
        }


        //Download to excel
        public IActionResult DownloadToExel()
        {
            using (var workbook = new XLWorkbook())
            {
                var worksheet = workbook.Worksheets.Add("_contacts");
                var currentRow = 1;
                worksheet.Cell(currentRow, 1).Value = "Id";
                worksheet.Cell(currentRow, 2).Value = "First Name";
                worksheet.Cell(currentRow, 3).Value = "Last Name";
                worksheet.Cell(currentRow, 4).Value = "Phone Number";
                worksheet.Cell(currentRow, 5).Value = "Email";
                worksheet.Cell(currentRow, 6).Value = "Address";
                foreach (var contact in _contacts)
                {
                    currentRow++;
                    worksheet.Cell(currentRow, 1).Value = contact.ID;
                    worksheet.Cell(currentRow, 2).Value = contact.Name;
                    worksheet.Cell(currentRow, 3).Value = contact.Lastname;
                    worksheet.Cell(currentRow, 4).Value = contact.Phone;
                    worksheet.Cell(currentRow, 5).Value = contact.Email;
                    worksheet.Cell(currentRow, 6).Value = contact.Address;
                }
                using (var stream = new MemoryStream())
                {
                    workbook.SaveAs(stream);
                    var content = stream.ToArray();
                    return File(content, "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet", "search_results.xlsx");
                }
            }
        }


        [Authorize(Roles = "Admin, User")]
        // GET: Contacts/Details/5
        public async Task<IActionResult> Details(int? id)
        {
            if (id == null || _context.Contacts == null)
            {
                return NotFound();
            }

            var contacts = await _context.Contacts
                .FirstOrDefaultAsync(m => m.ID == id);
            if (contacts == null)
            {
                return NotFound();
            }

            return View(contacts);
        }


        [Authorize(Roles = "Admin")]
        // GET: Contacts/Create
        public IActionResult Create()
        {
            return View();
        }


        [Authorize(Roles = "Admin")]
        // POST: Contacts/Create
        [HttpPost]
        [ValidateAntiForgeryToken]
        public async Task<IActionResult> Create([Bind("ID,Name,Lastname,Email,Phone,Address")] Contacts contacts)
        {
            if (ModelState.IsValid)
            {
                _context.Add(contacts);
                await _context.SaveChangesAsync();
                return RedirectToAction(nameof(Index));
            }
            return View(contacts);
        }


        [Authorize(Roles = "Admin")]
        // GET: Contacts/Edit/5
        public async Task<IActionResult> Edit(int? id)
        {
            if (id == null || _context.Contacts == null)
            {
                return NotFound();
            }

            var contacts = await _context.Contacts.FindAsync(id);
            if (contacts == null)
            {
                return NotFound();
            }
            return View(contacts);
        }

        [Authorize(Roles = "Admin")]
        // POST: Contacts/Edit/5
        [HttpPost]
        [ValidateAntiForgeryToken]
        public async Task<IActionResult> Edit(int id, [Bind("ID,Name,Lastname,Email,Phone,Address")] Contacts contacts)
        {
            if (id != contacts.ID)
            {
                return NotFound();
            }

            if (ModelState.IsValid)
            {
                try
                {
                    _context.Update(contacts);
                    await _context.SaveChangesAsync();
                }
                catch (DbUpdateConcurrencyException)
                {
                    if (!ContactsExists(contacts.ID))
                    {
                        return NotFound();
                    }
                    else
                    {
                        throw;
                    }
                }
                return RedirectToAction(nameof(Index));
            }
            return View(contacts);
        }


        [Authorize(Roles = "Admin")]
        // GET: Contacts/Delete/5
        public async Task<IActionResult> Delete(int? id)
        {
            if (id == null || _context.Contacts == null)
            {
                return NotFound();
            }

            var contacts = await _context.Contacts
                .FirstOrDefaultAsync(m => m.ID == id);
            if (contacts == null)
            {
                return NotFound();
            }

            return View(contacts);
        }


        [Authorize(Roles = "Admin")]
        // POST: Contacts/Delete/5
        [HttpPost, ActionName("Delete")]
        [ValidateAntiForgeryToken]
        public async Task<IActionResult> DeleteConfirmed(int id)
        {
            if (_context.Contacts == null)
            {
                return Problem("Entity set 'ApplicationDbContext.Contacts'  is null.");
            }
            var contacts = await _context.Contacts.FindAsync(id);
            if (contacts != null)
            {
                _context.Contacts.Remove(contacts);
            }
            
            await _context.SaveChangesAsync();
            return RedirectToAction(nameof(Index));
        }

        private bool ContactsExists(int id)
        {
          return _context.Contacts.Any(e => e.ID == id);
        }
    }
}
