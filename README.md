# City Year Automation Library

The City Year Automation library is an effort to create a toolbox of automation scripts for solving various problems faced
by Impact Analytics points. Users of this library will need some basic understanding of Python, but the design of the
library is such that most of the coding is already done in some wrapper functions.  An IA point trying to implement these
tools should typically just need to write some small functions for their specific use case and perhaps load some data into a
specified format in Excel.

## Packages

* `cyschoolhouse`
  * An adapter for cyschoolhouse (Salesforce) that supports automated actions and database queries. This package is largely a set of helper functions built around the [simple-salesforce](https://github.com/simple-salesforce/simple-salesforce) package. Currently supports section creation, student uploads, syncing student enrollment across multiple sections, and sending email.
* `excel-updater`
  * A tool for updating Excel Workbooks. Can currently update excel workbooks, handle sheet protection and hiding, and provides a structure for writing functions to update specific workbooks in a particular order.
* `selenium-testing` (testing only)
  * This folder is largely a set of testing scripts used as a proof of concept of a few different features of Selenium, and is only relevant if you're interested in some of the more advanced features that are being tested.

## Set-up (Windows)

1. Install Python 3.8. If you are new to Python, [see this guide](README-setup-python.md).
2. [Install GitHub Desktop](https://desktop.github.com/) and clone this repository to your computer.
3. In your console, create a virtual environment and activate it.
    ```console
    python -m venv env
    ```
    ```console
    env\Scripts\activate
    ```
4. Update `wheel` (required to install `xlwings`, see [this issue](https://github.com/xlwings/xlwings/issues/1243))
    ```console
    pip install --upgrade wheel
    ```
5. Install all the third party Python packages that are required for this project.
    ```console
    pip install -r requirements.txt
    ```
6. Copy `.env.sample` to `.env`.
    ```console
    copy .env.sample .env
    ```
7.  Fill in the details of `.env`. Do not edit `.env.sample`. Only edit `.env`. Saleforce credentials can be found on Salesforce as below:
    * `SF_USER`: Found at `My Settings > Personal > Personal Information`. Look for the `Username` field. This will be in the form `xxxxxxx@cityyear.org.cyschorgb`.
    * `SF_PASS`: Your password might not be the same as Okta. You can reset your Salesforce password under `My Settings > Personal > Change My Password`. This will not affect your Okta single sign on.
    * `SF_TOKEN`: Under `My Settings > Personal > Personal Information`, choose `Reset Security Token`. This will trigger an email to your inbox containing your security token.
8. [Install Firefox](https://www.mozilla.org/en-US/firefox/new/). This is the browser used for Salesforce automation.
9. Geckodriver is a tool used to automate tasks in Firefox. This driver is provided in this project at `./geckodriver/geckodriver.exe`.
If you don't trust this executable, you can replace it with the version [provided here](https://github.com/mozilla/geckodriver/releases).

Some scripts are used to manipulate files in cyconnect (SharePoint). This requires that the user map SharePoint as a network drive. Follow [this visual guide](README-setup-cyc.md) to set it up.

## cyschoolhouse Package

Files and folders mentioned in this section are relative to `./cyautomation/cyschoolhouse`.

* `cyschoolhousesuite.py`
  * A suite of wrapper functions for tasks common to anything involved in automating cyschoolhouse.  Allows user to call functions like `open_cyschoolhouse` instead of making direct calls to selenium.
* `section_creation.py`
  * The set of wrapper functions for creating sections.
* `service_trackers.py`
  * Generates pdf reports for each AmeriCorps Member on which they can manually track their weekly service implementation.
* `input_files` folder
  * Contains Excel workbooks that contain data to be uploaded. Theses are typically not used in Chicago.
* `templates` folder
  * Contains Excel workbooks that are populated by scripts and then written to cyconnect (SharePoint).

## Usage

For more examples on how Chicago uses these scripts in production, see `CHI-schedule.py` and `CHI-schedule-nb.ipynb`.

``` python
import cyautomation.cyschoolhouse as cysh

# Return a DataFrame of object records with all fields
cysh.get_object_df(object_name='Staff__c')

# Return a list of fields for a given object
cysh.get_object_fields(object_name='Staff__c')

# Return a DataFrame of object records with specific fields
cysh.get_object_df(
    object_name='Staff__c',
    field_list=['Individual__c', 'Name', 'CreatedDate', 'Reference_Id__c', 'Site__c', 'Role__c']
  )

# Return a DataFrame with a filter
cysh.get_object_df(
    object_name='Staff__c',
    field_list=['Individual__c', 'Name', 'Role__c'],
    where=f"Site__c = 'Chicago'"
)
```

## Contribute

The easiest way to get started is to dive into the code, and when you find something that doesn't make sense, post an issue.  If
you keep getting an error when you're running the code, post an issue.  This doesn't require any coding beyond trying to get the scripts to run.

If you want to contribute code to the project, follow the traditional [GitHub workflow](https://guides.github.com/introduction/flow/):
1. fork the repository
1. create a branch with a descriptive name
1. implement your code improvements
1. submit a pull request
