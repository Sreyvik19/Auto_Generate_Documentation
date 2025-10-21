# Hello From About Page


This is an **example documentation** using a modern style.

## Features

    Try the menu items and navigate back and forth between pages. Then click on Search. A search dialog will appear, allowing you to search for any text on any page. Notice that the search results include every occurrence of the search term on the site and links directly to the section of the page in which the search term appears. You get all of that with no effort or configuration on your part!

---
![Alt text](https://www.mkdocs.org/img/search.png)

- Easy to read
- Clean layout
- Custom colors

## Code Example

    Object-Oriented Programming (OOP) is a programming style that organizes code using “objects”, which combine data (attributes) and behavior (methods/functions). It makes code more structured, reusable, and easier to manage.

```python
# A simple library management system in Python

class Book:
    def __init__(self, title, author, year, available=True):
        self.title = title
        self.author = author
        self.year = year
        self.available = available

    def __str__(self):
        status = "Available" if self.available else "Checked out"
        return f"{self.title} by {self.author} ({self.year}) - {status}"

class Library:
    def __init__(self):
        self.books = []

    def add_book(self, book):
        self.books.append(book)
        print(f"Added: {book.title}")

    def list_books(self):
        print("Library Collection:")
        for book in self.books:
            print(book)

    def check_out(self, title):
        for book in self.books:
            if book.title == title:
                if book.available:
                    book.available = False
                    print(f"You have checked out: {book.title}")
                else:
                    print(f"Sorry, {book.title} is already checked out.")
                return
        print(f"Book titled '{title}' not found.")

    def return_book(self, title):
        for book in self.books:
            if book.title == title:
                if not book.available:
                    book.available = True
                    print(f"You have returned: {book.title}")
                else:
                    print(f"{book.title} was not checked out.")
                return
        print(f"Book titled '{title}' not found.")


# Example usage
library = Library()

book1 = Book("1984", "George Orwell", 1949)
book2 = Book("To Kill a Mockingbird", "Harper Lee", 1960)
book3 = Book("The Great Gatsby", "F. Scott Fitzgerald", 1925)

library.add_book(book1)
library.add_book(book2)
library.add_book(book3)

library.list_books()
library.check_out("1984")
library.check_out("1984")
library.return_book("1984")
library.list_books()
library.check_out("The Catcher in the Rye")  # Not in library

```
# Getting Help

See the [User Guide](https://www.mkdocs.org/user-guide/) for more complete documentation of all of MkDocs' features.

To get help with MkDocs, please use the [GitHub Discussions](https://github.com/mkdocs/mkdocs/discussions) or [GitHub Issues](https://github.com/mkdocs/mkdocs/issues).

---
