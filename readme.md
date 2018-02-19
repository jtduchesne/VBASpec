# VBASpec
Test-Driven Development (TDD) platform for VBA.

Strongly inspired by *Ruby*'s [RSpec](https://github.com/rspec/rspec-core).  
Also inspired by [VBA-TDD](https://github.com/VBA-tools/VBA-TDD) for the "`With` structure" and the [*Expectation*](#expectations) syntax.

## Installation

1. [Download the file `VBASpec.xlam`.](https://github.com/jtduchesne/VBASpec/raw/master/VBASpec.xlam)
2. Put it in your Office Addins folder (usually `~\AppData\Roaming\Microsoft\Addins\`).
3. [Follow this guide](https://support.office.com/en-us/article/add-or-remove-add-ins-in-excel-0af570c4-5cf3-4fa9-9b88-403625a0b460) to activate it via the *Addins* dialog.
4. [And this one](https://msdn.microsoft.com/en-us/vba/language-reference-vba/articles/check-or-add-an-object-library-reference) to add the Reference in your code.

## Usage

First create an instance:
```vb
Dim Suite As VBASpecSuite
Set Suite = New VBASpecSuite
```
Then _Describe_ what you are testing:
```vb
With Suite.Describe("Something")
    ...
```
Then explain what _It_ is doing:
```vb
With Suite.Describe("Something")
    With .It("exists")
        ...
```
Then code what is <i>Expect</i>ed:
```vb
With Suite.Describe("Something")
    With .It("exists")
        .Expect(something).ToNotBeNothing
```
Under the hood, the `Describe()` method returns an *ExampleGroup* class which is passed to the `With` statement. You can then access its `It()` method (notice the preceding `.`), which then returns an *Example* class which have a `Expect()` method (again the preceding `.`) that returns an [*Expectation*](#expectations) class.  

So the last piece of code could as well have been written like so:
```vb
Suite.Describe("Something").It("exists").Expect(something).ToNotBeNothing
```
and it would give the exact same result.

But when using the `With` syntax, we can chain as many sibling methods as we like *inside* the same parent:
```vb
With Suite.Describe("Something")
    With .It("exists")
        .Expect(something).ToNotBeNothing
    End With
    With .It("is a number, not a string")
        .Expect(something).ToBeAn "Integer"
        .Expect(something).ToNotBeA "String"
    End With
End With
```
which mimics the *closure syntax*, found in [interpreted languages](https://en.wikipedia.org/wiki/Interpreted_language), that would otherwise not be possible in a [compiled language](https://en.wikipedia.org/wiki/Compiled_language) like VBA.

## Expectations:

[`VBASpecExpectation`](https://github.com/jtduchesne/VBASpec/blob/master/VBASpecExpectation.cls) class provides methods that lets you express expected outcomes on an object in an example.
```vb
.Expect(1 + 1).ToEqual 2  'Passes
.Expect(1 + 1).ToEqual 3  'Fails with message: "Expected 2 to equal 3"
```

Here is the list of all expectations, along with their negated counterpart:

#### Equivalence
```vb
.Expect(Actual).ToEqual Expected    'Actual = Expected
.Expect(Actual).ToNotEqual Expected
.Expect(Actual).ToBe Expected       'Actual Is Expected
.Expect(Actual).ToNotBe Expected
```
#### Types
```vb
.Expect(Actual).ToBeA/ToBeAn Expected       'TypeName(Actual) = Expected
.Expect(Actual).ToNotBeA/ToNotBeAn Expected
```
#### Emptyness
```vb
.Expect(Actual).ToBeEmpty           'Actual = (Empty|Nothing|Null|Missing|"")
.Expect(Actual).ToNotBeEmpty
.Expect(Actual).ToBeNothing         'Actual Is Nothing
.Expect(Actual).ToNotBeNothing
.Expect(Actual).ToBeZero            'Actual is equal to 0
.Expect(Actual).ToNotBeZero
```
#### Truthyness
```vb
.Expect(Actual).ToBeTrue            'Actual = True
.Expect(Actual).ToBeTruthy          'Actual evaluates to True
.Expect(Actual).ToBeFalse           'Actual = False
.Expect(Actual).ToBeFalsy           'Actual evaluates to False
```
#### Comparisons
```vb
.Expect(Actual).ToBeLessThan/ToBeLT Expected
.Expect(Actual).ToBeLessThanOrEqual/ToBeLTE Expected
.Expect(Actual).ToBeGreaterThan/ToBeGT Expected
.Expect(Actual).ToBeGreaterThanOrEqual/ToBeGTE Expected

.Expect(Actual).ToBeCloseTo Expected, [SignificantFigures (Default: 2)]
.Expect(Actual).ToNotBeCloseTo Expected, [SignificantFigures (Default: 2)]
```
#### Collection/String membership
```vb
.Expect(Actual).ToInclude Expected      'Expected is included in Actual
.Expect(Actual).ToNotInclude Expected

.Expect(Actual).ToBeginWith Expected    'Expected is at the beginning of Actual
.Expect(Actual).ToNotBeginWith Expected
.Expect(Actual).ToEndWith Expected      'Expected is at the end of Actual
.Expect(Actual).ToNotEndWith Expected
                                        '(These all work with Arrays, Collections and Strings)
```
#### Expecting errors
```vb
On Error Resume Next '<- Really important
var = 1 / 0          '   The expectation must be placed AFTER the line that cause the error
.Expect.Error           'Passes when ANY error is raised
.Expect.Error(11)       'Passes, since 'Division by zero (Error 11)' was raised
.Expect.Error(12)       'Fails
On Error Goto 0
```
#### Expecting no errors
```vb
On Error Resume Next '<- Really important
var = 1 / 2          '   The expectation must be placed AFTER the line that could cause the error
.Expect.NoError 'Passes
var = 1 / 0
.Expect.NoError 'Fails
On Error Goto 0
```
```vb
.Expect.NoError.WasRaised   'The function .WasRaised can also be added at the end
.Expect.Error(11).WasRaised 'Mostly for aesthetic reasons...
.Expect.Error.WasRaised
```

## Example:

_modTest.bas_ :
```vb
Sub Test()
    Dim Suite as VBASpecSuite
    Set Suite = New VBASpecSuite
    
    With Suite.Describe("modUtils")
        With .Describe(".Max([ParamArray])")
            With .It("returns the highest number")
                .Expect(Max(2, 3, 5, 8)).ToEqual 8
                .Expect(Max(8, 5, 3, 2)).ToEqual 8
            End With
            
            With .It("works with decimal numbers")
                .Expect(Max(2.3, 5.8, 13.21)).ToEqual 13.21
            End With
        End With
    End With
End Sub
```
_modUtils.bas_ :
```vb
Public Function Max(ParamArray Numbers())
    Dim Number As Variant
    For Each Number In Numbers
        If Number > Max Then Max = Number
    Next Number
End Function
```
