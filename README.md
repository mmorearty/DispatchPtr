# IDispatch wrapper

_**Note:** The following was written in the year 2000._

ATL, MFC, and the Visual C++ runtime library already have wrappers for
`IDispatch` (`CComDispatchDriver`, `COleDispatchDriver`, and the
`_com_dispatch_...` methods, respectively). However, all three of these suffer
from syntax that is much more awkward than what you can write in languages such
as Visual Basic and Javascript. This wrapper, `CDispatchPtr` lets you write
code like this:

```c++
    CDispatchPtr htmldoc = ...;

    _bstr_t html = htmldoc.Get("body").Get("innerHTML");
    htmldoc.Set("title", "New Title");
    htmldoc.Get("body").Get("firstChild").Invoke(
        "insertAdjacentText", "afterBegin", "hello world");
```

## Methods

All of the methods of `CDispatchPtr` take a first argument that specifies which
member (i.e. property or method) of the `IDispatch` is being accessed. You can
pass _any one_ of the following three data types for this parameter:

*   A regular ANSI string, e.g. `"name"`
*   A UNICODE string, e.g. `L"name"`
*   A `DISPID`, e.g. `DISPID_NAME`

Using a `DISPID` is most efficient, but least convenient (because you have to
figure out what the `DISPID` is by muddling through documentation and header
files). Using a UNICODE string is second-best. Using an ANSI string is least
efficient. However, the efficiency difference between using a UNICODE string
and an ANSI string is very small, so don't be afraid to use ANSI strings if you
want to.

`CDispatchPtr` has the following methods:

**Get(property)**

Returns the value of the property as a `_variant_t`. `_variant_t` supports
implicit casting to a number of common types, so all of the following are
legal:  

```c++
    long length = myDispatch.Get("length");  
    _bstr_t title = myDispatch.Get("title");  
    CDispatchPtr body = myDispatch.Get("body");  
    CDispatchPtr htmlElement = myDispatch.Get("body").Get("firstChild");
```

**Set(property, \_variant\_t value)   <br>
SetRef(property, \_variant\_t value)**

Sets the value of a property. \_variant\_t supports implicit casting from a
number of common types, so all of the following are legal:  

```c++
    myDispatch.Set("length", 3L);  
    myDispatch.Set("title", "New Title");  
    myDispatch.SetRef("body", (IDispatch*) mybody);
```

**Invoke(method, ...)**

Invokes a method and returns its value. Due to the implementation, this is not
a true "varargs" function, but the implementation allows for up to 9 arguments.
Examples:  

```c++
    myDispatch.Invoke("onclick");  
    myDispatch.Invoke("insertAdjacentText", "afterBegin", "hello world");  
    long difference = myDispatch.Invoke("difference", 3L, 4L);
```

## Set() vs. SetRef()

The difference between Set() and SetRef() is subtle and annoying. If you know
Visual Basic, then the easiest way to explain it is that SetRef() is used for
VB's "Set" keyword, e.g. "Set x = y", and Set() is used for regular value
assignment, e.g. "x = y".

SetRef() sets the value of a property to be a _reference_ to the passed-in
value. This only makes sense if the passed-in value is of type `(IDispatch*)` or
`(IUnknown*)`.

In most cases, you'll just use `Set()`.

## CDispatchPtr, IDispatchPtr, and Exceptions

`CDispatchPtr` is a subclass of `IDispatchPtr`, with the Get/Set/SetRef/Invoke
methods added. `IDispatchPtr` is `_com_ptr_t` templated on `IDispatch`.

Because of this, any place in your code where you would have used an
`IDispatchPtr`, you can use a `CDispatchPtr` in its place.

You need to be aware that, since `_com_ptr_t` throws exceptions when errors
occur, `CDispatchPtr` does as well. I chose to subclass `IDispatchPtr` instead of
ATL's `CComPtr` because the `Get()` and `Invoke()` functions, by necessity of what
they're for, can't return an error code as their return value, so they must
throw exceptions to indicate errors. `CComPtr` does not throw exceptions, but
`_com_ptr_t` does.

In addition to those exceptions that result from `_com_ptr_t`'s methods being
called, `CDispatchPtr` will also raise an exception if an `IDispatch`-based method
or property that you try to invoke returns an error code -- e.g. because you
passed invalid arguments or because the implementing object fails -- an
exception (of type `_com_error`) will be thrown.

Because of this, you will need to use try/catch blocks in your code, just as
you would if you used `IDispatchPtr` or any other `_com_ptr_t`-derived type.

`_com_error` has a member function called ErrorMessage() which returns the text
error message for the HRESULT that caused the error, so for quickie debugging
code, this is a good starting point:

```c++
    try
    {
        // ... your code here ...
    }
    catch( _com_error&amp; err )
    {
        printf( "Error: %sn", err.ErrorMessage() );
    }
```

## Performance

Methods invoked via `IDispatch` are always somewhat slower than those invoked via
direct interfaces. However, the performance difference is small enough not to
be a major concern in most cases.

And in the case of accessing properties and methods of MSHTML, going through
`IDispatch` is dramatically easier than going through direct interfaces, and thus
is almost always worth the small loss in performance. For example, the
Javascript expression

```c++
    document.body.firstChild
```

is a real pain to write in C++ using direct interfaces:

```c++
    IHTMLDocument2Ptr document;
    ((MSHTML::IHTMLDOMNodePtr)document->body)->firstChild
```

Not only is that hard to read, but the hardest part is that the Microsoft
documentation is usually not too good about specifying which interface each
MSHTML property and method actually belongs to, so figuring out that firstChild
is a member of IHTMLDOMNodePtr is laborious (usually I find it by searching
through the #import-generated mshtml.tlh).

With `CDispatchPtr`, though, it's much easier to write, and also to read:

```c++
    CDispatchPtr document;
    document.Get("body").Get("firstChild")
```

I ran some crude timing tests to compare C++ native calls, C++ with
`CDispatchPtr` (and thus `IDispatch`), and Javascript (which also uses `IDispatch`).
In each case, the code repeated a loop 1000 times, and inside the loop, the
code did a number of Get/Set/Invoke operations on an HTML document.

Here are the results:

* C++ with native calls: 1.45 seconds
* C++ with `CDispatchPtr`: 2.33 seconds
* Javascript: 3.74 seconds

As you can see, `CDispatchPtr` is about 60% slower than using native calls.
That's quite a significant slowdown. However, as I mentioned, this was
repeating a loop 1000 times (on my PIII 500MHz). The inner loop performed 13
Get/Set/Invoke operations, for a total of 13,000 operations. Considering that
`CDispatchPtr` completed that in 2.33 seconds, it's clear that in most cases, the
extra overhead of `IDispatch` is not a problem.
