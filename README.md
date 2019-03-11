# port_node_to_odbc
Simple asynchronous wraper for ODBC conections
# requirements
To run this code on the node.js side you need install first node-edge library
``` 
npm install edge-js --save
```
# How to use
Create simple functions to establish connections '../connect/access.js'
``` node
const  edge   = require('edge-js');
const  util = require('util');
const  oledb  = edge.func({ assemblyFile: 'Class_odbc.dll', typeName: 'Class_odbc.Startup', methodName: 'Invoke' });
function get_ODBC(options) {
  return new Promise((resolve, reject) => {
    oledb(options, (error,result) => {
      if(error) {
        reject(error);}
      else {
        try {
          resolve(result);
        } catch (e) {
          reject({ message: e.message });
        }
      }
    });
  });
}

async function get(options) {
  var data ={}
  const result=await get_ODBC(options);
  data.valid = true;
  data.records = result;
  return data
}

module.exports = {
  Getoledb: (options) => get(options)
};

```
In the rest of program Call
``` node
const db = require('../connect/access.js')
const Adsn="Provider=Microsoft.ace.OLEDB.12.0;Data Source=user_access.accdb;Jet OLEDB:Database Password=password;";

const options = {
      dsn :  Adsn,
      query : "INSERT INTO USERS ( login, salt, verifier, email ) values ( @login, @salt, @verifier, @email )",
      prepare: "true",
      Values : {
        Val_name1: '@login',
        Value1:username,
        Type1:'VarWChar',
        Len1:124,
        Val_name2: '@salt',
        Value2:salt,
        Type2:'VarWChar',
        Len2:4048,
        Val_name3: '@verifier',
        Value3:verifier,
        Type3:'VarWChar',
        Len3:6024,
        Val_name4: '@email',
        Value4:email,
        Type4:'VarWChar',
        Len4:124,
        }
      };
      const result = await db.Getoledb(options);
      
 options = {
    dsn :  Adsn,
    query : "Select count(login) as exist from USERS where login=@login",
    prepare: "true",
    Values : {
       Val_name1: '@login',
       Value1:username,
       Type1:'VarWChar',
       Len1:50,
      }
  };
  result = await db.Getoledb(options);
