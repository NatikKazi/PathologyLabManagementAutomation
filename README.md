# PathologyLabManagementAutomation

This repository contains automation test scripts for a Pathology Lab Management web application using Selenium, C#, NUnit, and data-driven testing. The tests are based on the Page Object Model (POM) design pattern.

## Application URL

The application under test can be accessed at [https://gor-pathology.web.app/patients/add](https://gor-pathology.web.app/patients/add).

## Topics Covered

The automation test scripts cover the following topics and functionalities of the Pathology Lab Management application:

- **Login Page**: Automates the login process with test data read from an Excel file.

- **Home Page**: Automates interactions with the home page, including entering cost and discount values, checking for alerts, and clearing inputs.

- **Patient Page**: Automates the patient creation process, including entering patient information, general details, and adding tests. It also verifies the created patient in the list.

## Project Structure

The project is organized into several folders:

- `Drivers`: Contains the base driver setup for Selenium.

- `Models`: Contains model classes used for data serialization. JSON is used for data storage and retrieval.

- `Pages`: Contains page object classes for different application pages (Login, Home, Patient).

- `Tests`: Contains test classes that use the page objects to perform automation.

## Testing Methodologies

- **Page Object Model (POM)**: The project follows the POM design pattern to maintain clean and organized test scripts.

- **Data-Driven Testing**: Test data is read from Excel files located in the `DataTest` folder, enabling multiple test scenarios with different data sets.

- **C# with Selenium and NUnit**: The tests are written in C# using the Selenium WebDriver for web automation and NUnit as the testing framework.

- **JSON Usage**: JSON is used for data serialization and deserialization. It helps manage and transfer data between test scripts and the application.

## Usage

To run the automation tests:

1. Clone this repository to your local machine.

2. Install the necessary dependencies. Make sure you have Chrome WebDriver set up.

3. Open the solution in Visual Studio or your preferred IDE.

4. Configure your test data in Excel files located in the `DataTest` folder.

5. Run the test classes located in the `Tests` folder.

## Dependencies

- Selenium WebDriver
- WebDriverManager
- Spire.Xls
- NUnit

## Contributors

- Natik Kazi

