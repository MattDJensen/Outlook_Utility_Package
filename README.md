## Outlook Msg Automation functions

#### display_email: Display draft of email
```python
def display_email(message: str, subject: str, to_list: str, cc_list: str):
    """
        :param message: HTML String with Email message contained. See Examples/Email_Strings.py
        :param subject: Subject String
        :param to_list: Semicolon separated list of email addresses. (ex - a@abc.com; b@abc.com; c@abc.com;)
        :param cc_list: Semicolon separated list of email addresses. (ex - a@abc.com; b@abc.com; c@abc.com;)
        """
```
##### Example Call
```python
from outlookutility import display_email
test_html = f"""
    <HTML>
    <BODY>
    Package Testing Email
    <br>
    </BODY>
    </HTML>"""

display_email(
    test_html,
    "PyPi Test",
    "a@abc.com; b@abc.com;",
    "c@abc.com;",
)
```

#### display_email_with_attachments: Display draft of email with attachments. Can send any number/type of attachments in email. 
```python
def display_email_with_attachments(message: str, subject: str, to_list: str, cc_list: str, *args):
    """
        :param message: HTML String with Email message contained. See Examples/Email_Body.html.
        :param subject: Subject String
        :param to_list: Semicolon separated list of email addresses. (ex - a@abc.com; b@abc.com; c@abc.com;)
        :param cc_list: Semicolon separated list of email addresses. (ex - a@abc.com; b@abc.com; c@abc.com;)
        :param args: Optional attachment arguments, pass as raw file path or stringified file path.
        """
```
##### Example Call
```python
from outlookutility import display_email_with_attachments
test_html = f"""
    <HTML>
    <BODY>
    Package testing email with attachments
    <br>
    </BODY>
    </HTML>"""

display_email_with_attachments(
    test_html,
    "PyPi Test",
    "a@abc.com; b@abc.com;",
    "c@abc.com;",
    r"C:\Users\user\test_1.txt",
    r"C:\Users\user\test_2.txt",
)
```

#### email_without_attachment: Send email without attachments. 
```python
def email_without_attachment(message: str, subject: str, to_list: str, cc_list: str):
    """
        :param message: HTML String with Email message contained. See Examples/Email_Strings.py
        :param subject: Subject String
        :param to_list: Semicolon separated list of email addresses. (ex - a@abc.com; b@abc.com; c@abc.com;)
        :param cc_list: Semicolon separated list of email addresses. (ex - a@abc.com; b@abc.com; c@abc.com;)
        """
```
##### Example Call
```python
from outlookutility import email_without_attachment
test_html = f"""
    <HTML>
    <BODY>
    Package Testing Email
    <br>
    </BODY>
    </HTML>"""

email_without_attachment(
    test_html,
    "PyPi Test",
    "a@abc.com; b@abc.com;",
    "c@abc.com;",
)
```

#### email_with_attachments: Send email with attachments. Can send any number/type of attachments in email. 
```python
def email_with_attachments(message: str, subject: str, to_list: str, cc_list: str, *args):
    """
        :param message: HTML String with Email message contained. See Examples/Email_Body.html.
        :param subject: Subject String
        :param to_list: Semicolon separated list of email addresses. (ex - a@abc.com; b@abc.com; c@abc.com;)
        :param cc_list: Semicolon separated list of email addresses. (ex - a@abc.com; b@abc.com; c@abc.com;)
        :param args: Optional attachment arguments, pass as raw file path or stringified file path.
        """
```
##### Example Call
```python
from outlookutility import email_with_attachments
test_html = f"""
    <HTML>
    <BODY>
    Package testing email with attachments
    <br>
    </BODY>
    </HTML>"""

email_with_attachments(
    test_html,
    "PyPi Test",
    "a@abc.com; b@abc.com;",
    "c@abc.com;",
    r"C:\Users\user\test_1.txt",
    r"C:\Users\user\test_2.txt",
)
```

#### notify_error: Automated email report for use in exception catch. 
```python
def notify_error(report_name, error_log, to_list: str):
    """

    :param to_list: List of emails to receive notification.
    :param report_name: Name of automated report.
    :param error_log: Raised exception or other error to report.
    """
```
##### Example Call
```python
from outlookutility import notify_error
import os
def foo():
    raise Exception('Error!')
try:
    foo()
except Exception as e:
    notify_error(f"{os.path.basename(__file__)}", e, "a@email.com")
```



#### default_table_style : Apply formatting to Pandas dataframe for use in email
```python
def default_table_style(df, index: False):
    """ Apply a default clean table style to pandas df.to_html() for use in email strings.

    :param index: Determines whether you want index displayed in the HTML. Defaults to False.
    :type index: Boolean
    :param df: Dataframe to apply the style to.
    :type df: Pandas Dataframe
    :return: HTML string for insertion in email.
    :rtype: string
    """
```
##### Example Call
```python
from outlookutility import default_table_style
import pandas as pd 
import numpy as np
df = pd.DataFrame(np.random.randint(0,100,size=(15, 4)), columns=list('ABCD'))
test_message = f"""
<HTML>
    <BODY>
     Email Text Here
     <br>
     {default_table_style(df,index=False)}
     <br>
    </BODY>
</HTML>
"""

```
### Example Output
<HTML>
    <BODY>
    <p>Email Text Here</p>
    <head>
        <style>
            table.dataframe {
                border-collapse: collapse;
            }
            table.dataframe th{
                border: 1px solid black;
                padding: 10px;
                font-size:14px;
                text-align: center;
            }
            table.dataframe td {
                border: 1px solid black;
                padding: 10px;
                font-size:12px;
                text-align: center;
            }
        }
        </style>
    </head>
    <table border="1" class="dataframe">
  <thead>
    <tr style="text-align: right;">
      <th>A</th>
      <th>B</th>
      <th>C</th>
      <th>D</th>
    </tr>
  </thead>
  <tbody>
    <tr>
      <td>28</td>
      <td>37</td>
      <td>9</td>
      <td>3</td>
    </tr>
    <tr>
      <td>38</td>
      <td>21</td>
      <td>72</td>
      <td>86</td>
    </tr>
    <tr>
      <td>54</td>
      <td>42</td>
      <td>52</td>
      <td>39</td>
    </tr>
    <tr>
      <td>54</td>
      <td>41</td>
      <td>94</td>
      <td>80</td>
    </tr>
    <tr>
      <td>99</td>
      <td>37</td>
      <td>6</td>
      <td>62</td>
    </tr>
    <tr>
      <td>58</td>
      <td>7</td>
      <td>72</td>
      <td>62</td>
    </tr>
    <tr>
      <td>46</td>
      <td>85</td>
      <td>49</td>
      <td>47</td>
    </tr>
    <tr>
      <td>66</td>
      <td>29</td>
      <td>36</td>
      <td>13</td>
    </tr>
    <tr>
      <td>74</td>
      <td>88</td>
      <td>40</td>
      <td>6</td>
    </tr>
    <tr>
      <td>98</td>
      <td>96</td>
      <td>59</td>
      <td>57</td>
    </tr>
    <tr>
      <td>77</td>
      <td>53</td>
      <td>10</td>
      <td>82</td>
    </tr>
    <tr>
      <td>68</td>
      <td>26</td>
      <td>87</td>
      <td>33</td>
    </tr>
    <tr>
      <td>95</td>
      <td>16</td>
      <td>79</td>
      <td>16</td>
    </tr>
    <tr>
      <td>80</td>
      <td>0</td>
      <td>39</td>
      <td>17</td>
    </tr>
    <tr>
      <td>80</td>
      <td>14</td>
      <td>50</td>
      <td>69</td>
    </tr>
  </tbody>
</table>
     <br>
    </BODY>
</HTML>

#### multi_table_style : Apply formatting to multiple Pandas dataframes for use in email
```python
def multi_table_style(df_list, index: False):
    """ Apply a default clean table style to pandas df.to_html() for use in email strings.
    This version returns multiple tables stacked on top of each other with a line break inbetween.

    :param index: Determines whether you want index displayed in the HTML. Defaults to False.
    :type index: Boolean
    :param df_list: List of dataframes to return in html format.
    :type df: Pandas Dataframe
    :return: HTML string for insertion in email.
    :rtype: string
    """
```
##### Example Call
```python
from outlookutility import multi_table_style
import pandas as pd 
import numpy as np
df = pd.DataFrame(np.random.randint(0,100,size=(15, 4)), columns=list('ABCD'))
df_list = [df,df]
formatted_tables = multi_table_style(df_list,index=False)
test_message = f"""
<HTML>
    <BODY>
     Email Text Here
     <br>
     {formatted_tables}
     <br>
    </BODY>
</HTML>
"""

```
### Example Output
<HTML>
    <BODY>
     <p> Email Text Here </p>
<head>
    <style>
        table.dataframe {
            border-collapse: collapse;
        }
        table.dataframe th{
            border: 1px solid black;
            padding: 10px;
            font-size:14px;
            text-align: center;
        }
        table.dataframe td {
            border: 1px solid black;
            padding: 10px;
            font-size:12px;
            text-align: center;
        }
    }
    </style>
</head>
<table border="1" class="dataframe">
  <thead>
    <tr style="text-align: right;">
      <th>A</th>
      <th>B</th>
      <th>C</th>
      <th>D</th>
    </tr>
  </thead>
  <tbody>
    <tr>
      <td>28</td>
      <td>32</td>
      <td>58</td>
      <td>45</td>
    </tr>
    <tr>
      <td>1</td>
      <td>37</td>
      <td>89</td>
      <td>69</td>
    </tr>
    <tr>
      <td>30</td>
      <td>12</td>
      <td>21</td>
      <td>60</td>
    </tr>
    <tr>
      <td>52</td>
      <td>26</td>
      <td>86</td>
      <td>44</td>
    </tr>
    <tr>
      <td>97</td>
      <td>62</td>
      <td>46</td>
      <td>32</td>
    </tr>
    <tr>
      <td>32</td>
      <td>8</td>
      <td>63</td>
      <td>35</td>
    </tr>
    <tr>
      <td>35</td>
      <td>8</td>
      <td>71</td>
      <td>89</td>
    </tr>
    <tr>
      <td>54</td>
      <td>14</td>
      <td>88</td>
      <td>65</td>
    </tr>
    <tr>
      <td>66</td>
      <td>11</td>
      <td>69</td>
      <td>58</td>
    </tr>
    <tr>
      <td>46</td>
      <td>84</td>
      <td>12</td>
      <td>79</td>
    </tr>
    <tr>
      <td>49</td>
      <td>60</td>
      <td>77</td>
      <td>88</td>
    </tr>
    <tr>
      <td>32</td>
      <td>22</td>
      <td>97</td>
      <td>50</td>
    </tr>
    <tr>
      <td>89</td>
      <td>27</td>
      <td>26</td>
      <td>35</td>
    </tr>
    <tr>
      <td>64</td>
      <td>4</td>
      <td>62</td>
      <td>11</td>
    </tr>
    <tr>
      <td>85</td>
      <td>99</td>
      <td>71</td>
      <td>18</td>
    </tr>
  </tbody>
</table><br/> <br/><table border="1" class="dataframe">
  <thead>
    <tr style="text-align: right;">
      <th>A</th>
      <th>B</th>
      <th>C</th>
      <th>D</th>
    </tr>
  </thead>
  <tbody>
    <tr>
      <td>28</td>
      <td>32</td>
      <td>58</td>
      <td>45</td>
    </tr>
    <tr>
      <td>1</td>
      <td>37</td>
      <td>89</td>
      <td>69</td>
    </tr>
    <tr>
      <td>30</td>
      <td>12</td>
      <td>21</td>
      <td>60</td>
    </tr>
    <tr>
      <td>52</td>
      <td>26</td>
      <td>86</td>
      <td>44</td>
    </tr>
    <tr>
      <td>97</td>
      <td>62</td>
      <td>46</td>
      <td>32</td>
    </tr>
    <tr>
      <td>32</td>
      <td>8</td>
      <td>63</td>
      <td>35</td>
    </tr>
    <tr>
      <td>35</td>
      <td>8</td>
      <td>71</td>
      <td>89</td>
    </tr>
    <tr>
      <td>54</td>
      <td>14</td>
      <td>88</td>
      <td>65</td>
    </tr>
    <tr>
      <td>66</td>
      <td>11</td>
      <td>69</td>
      <td>58</td>
    </tr>
    <tr>
      <td>46</td>
      <td>84</td>
      <td>12</td>
      <td>79</td>
    </tr>
    <tr>
      <td>49</td>
      <td>60</td>
      <td>77</td>
      <td>88</td>
    </tr>
    <tr>
      <td>32</td>
      <td>22</td>
      <td>97</td>
      <td>50</td>
    </tr>
    <tr>
      <td>89</td>
      <td>27</td>
      <td>26</td>
      <td>35</td>
    </tr>
    <tr>
      <td>64</td>
      <td>4</td>
      <td>62</td>
      <td>11</td>
    </tr>
    <tr>
      <td>85</td>
      <td>99</td>
      <td>71</td>
      <td>18</td>
    </tr>
  </tbody>
</table><br/> <br/>
     <br>
    </BODY>
</HTML>
