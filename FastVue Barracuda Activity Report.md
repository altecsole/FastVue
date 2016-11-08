# Using FastVue Reporter to create the Activity Report

  * Open FastVue Reporter and go to Reports -> Activity Report
  * Select the following options:

| Field | Operator | Values |
| --- | --- | --- |
| Referrer with Query | Contains | youtube.com/watch youtube.com/watch?v= |
| Origin Domain | Equal to | youtube.com |

  * Set your dates and times and then Run Report

Once the report has run, choose to export as CSV. You should then be able to use the file with the Powershell script to get the titles of the YouTube vidoes.

