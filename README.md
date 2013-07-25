# gs-spreadsheet-manager

A slightly more pleasant way of dealing with Google Spreadsheets in Google Apps Script.

## gs-managed-spreadsheet

An attempt to make the interface a little more object-oriented and ORM-y.

### How to use

Let's say you have a spreadsheet with a worksheet called `animals`. It looks like this.

<table>
  <tr>
    <th>name</th><th>description</th><th>count</th>
  </tr>
  <tr>
    <td>Koala</td><td>They eat a lot of eucalyptus</td><td>55</td>
  </tr>
  <tr>
    <td>Tiger</td><td>They eat a lot of human beings</td><td>2</td>
  </tr>
</table>

    // Let's get the animals worksheet
    var spreadsheet = ManagedSpreadsheet('SPREADSHEET_KEY');
    var worksheet = spreadsheet.sheet("animals");
    
    // Let's grab the info on the tiger
    var tiger = worksheet.find('name', 'Tiger')
    // { name: 'Tiger', description: 'They eat a lot of human beings', count: 2 }
    
    // Let's update the tiger's data
    tiger.count = 7;
    var index = worksheet.rowIndex('name', 'Tiger');
    worksheet.update(rowIndex, tiger);
    
    // Let's add another animal
    worksheet.append({name: 'Weasel', description: 'Always weaseling around', count: 10});

    // We no longer need animals we have less than 5 of
    worksheet.deleteWhere( function(row) {
      return row.count < 5;
    })
    
    // Let's grab what's left
    var everything = worksheet.all()
    // [{name: 'Koala', description: 'They eat a lot of eucalyptus', count: 55}, 
    //  {name: 'Weasel', description: 'Always weaseling around', count: 10}]

Lots of other stuff in there, too. Caching and forking and activating and copying across spreadsheets. You can see a lot of it in action in `gs-distribution-manager`, actually.

## gs-distribution-manager

Since you can only have so many worksheets/cells per spreadsheet, sometimes you need to distribute your data across multiple spreadsheets. That's what gs-distribution-manager does.

Best use case I can think of so far is taking the input from a google form and generating worksheets out of it. You have a worksheet named `template` and DistributionManager copies it to where you want it, storing a bit of info about where it went.