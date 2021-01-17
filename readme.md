# AccessTestSuite

AccessTestSuite is a Microsoft Access add-in for creating and running unit tests of VBA code within an Access database project.

## Installation

Attach `AccessTestSuite.accdb` to your VBA project through the References dialog. Then, from the debug window, enter the command 

```
TestsInstall "tests"
```

The parameter `"tests"` is optional; if used, it will attach a prefix to the installed test tables and form so they don't get in your way. For example, `TestsInstall "zz"` will create tables prefixed zz_ so they alphabetize to the bottom of your object list.

## Usage

Create tests using the `TestItems` form. You can also create them directly in the `TestItem` table, but then you will lose the various drop-down menus that make this sort of thing easier.

### Test Notes
* Only `Public` VBA functions can be tested.
* The tests can use only up to three parameters.
* If tested functions return with an error code (`Err.Number <> 0`), the test will report as having failed.
* Access tables do not store whitespace, so you can use these tokens in parameters or expected results: {NULL}, {NULLSTRING}, {TRUE}, {FALSE}, {SPACE}, {CRLF}.
* You can use the format `Array(a,b,c)` to pass arrays (in a Variant) to a test. 
* If you want to test a function that returns an array, use the test type `code-array` and enter the expected test results as a pipe-delimited string (`"a|b|c"`).

### Running

Run tests from the `TestItems` form, as well, or by using the command

```
TestsRun "tests"
```

## Notes

When not actively working with tests, I recommend that you remove the reference to `AccessTestSuite.accdb` from your References.

## Contributing
Pull requests are welcome. For major changes, please open an issue first to discuss what you would like to change.

Please make sure to update tests as appropriate.

The source code has been backed up using [msaccess-vcs-integration](https://github.com/msaccess-vcs-integration/msaccess-vcs-integration). MS Access backups are finicky; if you're doing a pull request, using that package is the easiest way, I think.

## Development Pathway

This is a pretty rudimentary library; here are some of my current plans.

* add attaching/detaching reference to test suite manually from the test runner form
* add a way to return an object and test a property of that object
* allow result to return variant so we can test for NULL, etc.

## License
[MIT](https://choosealicense.com/licenses/mit/)