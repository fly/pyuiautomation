pyuiautomation
==============

Basic python wrapper for Microsoft UI Automation functionality.


Please note, I had created this several years ago, when I needed it and there were no other python modules that provide such functional.
It only implements functionality that I needed at that time.
Currently I don't do any automation and it is unlikely that I will do it in the future. So it's unlikely that I will continue development of this module.
AFAIK, [pywinauto](https://pywinauto.github.io/) introduced MS UI Automation support in the latest release. If you need such functionality, you probably should give it a try.

Example of usage (for Windows < 10):
```python
import subprocess
import time

import pyuiautomation

if __name__ == '__main__':
    subprocess.Popen('calc.exe')
    time.sleep(1)
    root_element = pyuiautomation.GetRootElement()
    print root_element
    calc = root_element.findfirst('descendants', Name='Calculator',
                                  ControlType=pyuiautomation.UIAutomationClient.UIA_WindowControlTypeId)
    print calc
    button_2 = calc.findfirst('descendants', Name='2')
    print button_2
    button_add = calc.findfirst('descendants', Name='Add')
    print button_add
    button_equals = calc.findfirst('descendants', Name='Equals')
    print button_equals
    button_2.Invoke()
    button_add.Invoke()
    button_2.Invoke()
    button_equals.Invoke()
    result = calc.findfirst('descendants', AutomationId='158').Name
    print '2 + 2 = %s' % result

```

Windows 10 has different program for calculator. The example code may be the following:
```python
button_2 = calc.findfirst('descendants', Name='Two')
button_add = calc.findfirst('descendants', Name='Plus')
button_equals = calc.findfirst('descendants', Name='Equals')
button_2.Invoke()
button_add.Invoke()
button_2.Invoke()
button_equals.Invoke()
result = calc.findfirst('descendants', AutomationId='CalculatorResults').Name
print result
```