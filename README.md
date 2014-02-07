pyuiautomation
==============

Python realisation of basic Microsoft UI Automation functionality

Please note, this one contais only functionality that I currently need. New functionality may be added occasionally, when I need it.

Example of usage:
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
