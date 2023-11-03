 const excel_file = document.getElementById('excel_file');

excel_file.addEventListener('change', (event) => {

      if (!['application/vnd.openxmlformats-officedocument.spreadsheetml.sheet', 'application/vnd.ms-excel'].includes(event.target.files[0].type)) {
            document.getElementById('excel_data').innerHTML = '<div class="alert alert-danger">Only .xlsx or .xls file format are allowed</div>';

            excel_file.value = '';

            return false;
      }

      var reader = new FileReader();

      reader.readAsArrayBuffer(event.target.files[0]);

      reader.onload = function (event) {

            var data = new Uint8Array(reader.result);

            var work_book = XLSX.read(data, { type: 'array' });

            var sheet_name = work_book.SheetNames;

            var sheet_data = XLSX.utils.sheet_to_json(work_book.Sheets[sheet_name[0]], { header: 1 });

            if (sheet_data.length > 0) {
                  var table_output = '';

                  for (var row = 0; row < sheet_data.length; row++) {

                        table_output += `<div class="page"> <div class="subpage"><table style="width:100%" class=" table table-borderless " id="actbl"> <tr><th style="width:4cm!important" class="" colspan="1"><img  id="llogo_${row}" style="width:4cm"src="https://mis.cept.gov.in/Library/Images/indiapostlogo.jpg"></th><th id="ofc_${row}"  class="" colspan="3">Department of Posts<br/>O/o Supdt. of Post Offices, Delhi Division, Delhi-110006<th class="" style="width:4" colspan="1"><img id="rlogo_${row}"  style="width:4cm" src="https://mis.cept.gov.in/Library/Images/right_side_logo.jpg"></th></tr><tr style="font-size: 20px;"><th colspan="5"><exm id="exam_${row}" >Deen Dayal Sparsh Scholarship Scheme Year-9999</exm><br/>
<ac>Admit Card</ac></th></tr>`;

                        for (var cell = 0; cell < sheet_data[row].length; cell++) {
                          var celllngth = sheet_data[row].length;
                          var rwspan = celllngth-8;

                              if (row ==0) {
  
                              if (cell<5)  { table_output += '<tr class=""><td colspan="1" class="' + '-cell' + cell + '">' + sheet_data[0][cell] + '</td>'+'<td colspan="2" class="' + '-cell' + cell + '">' + sheet_data[row][cell]+'</td>';

                                            if (cell==0){  table_output += '<td style="width:4cm;text-align:center" colspan="2" rowspan="8">Passport size photo</td></tr>'};

                                   // table_output += '<th class="row' + '-cell' + cell + '">' + sheet_data[row][cell] + '</th>';
}
                                else{
table_output += '<tr colspan="1" class=""><td class="' + '-cell' + cell + '">' + sheet_data[0][cell] + '</td>'+'<td colspan="2" class="' + '-cell' + cell + '">' + sheet_data[row][cell]+'</td>';
                                  if (cell>7){  table_output += '<td style="width:4cm;text-align:center" colspan="2" rowspan="'+rwspan+'"></td></tr>'};
}

                              } 
                          else 
                          {

                                     if (cell<5)  { table_output += '<tr class=""><td colspan="1" class="' + '-cell' + cell + '">' + sheet_data[0][cell] + '</td>'+'<td colspan="2" class="' + '-cell' + cell + '">' + sheet_data[row][cell]+'</td>';
if (cell==0){  table_output += '<td style="width:4cm;text-align:center" colspan="2" rowspan="8">Passport size photo</td></tr>'};

                                   // table_output += '<th class="row' + '-cell' + cell + '">' + sheet_data[row][cell] + '</th>';
}
                          else{
table_output += '<tr colspan="1" class=""><td class="' + '-cell' + cell + '">' + sheet_data[0][cell] + '</td>'+'<td colspan="2" class="' + '-cell' + cell + '">' + sheet_data[row][cell]+'</td>';
                                  if (cell>7){  table_output += '<td style="width:4cm;text-align:center" colspan="2" rowspan="'+rwspan+'"></td></tr>'};

                          }



                              }

                        }
//

                        table_output += `<tr class="">
<td colspan="4">Exam Controlling Authority</td>
<th style="text-align:center" ><img id="sign_${row}" height="60"><br/><ctrl id="ctrl_${row}">Supdt. of Post Offices<br/>Delhi Dn, Delhi-110006</ctrl></th>

</tr> 
<tr class="row">
<td colspan="5"></td>


</tr>

<tr> 
<th colspan="5">Instructions:</th><tr></tr>
<td id="ins_${row}" colspan="5">

</td></tr></table></div></div><br clear="all" style="page-break-before:always" />`;
                    
                    
                    
            

                  
                  }
                  table_output += '';
              
                    

                  document.getElementById('excel_data').innerHTML = table_output;
            for (var row = 0; row < sheet_data.length; row++) {  
              var ins=$("#text").val();
             var str= ins.replace(/(?:\r\n|\r|\n)/g, '<br>');  
             $(`#ins_${row}`).html(str);
            }
            
            }

            excel_file.value = '';

      }
  

}); 
  function updt()  { var rowCount = $('.page').length;
                   // alert(rowCount);
  for (var row = 0; row < rowCount; row++) {  
              var ins=$("#text").val();
    var exm=$("#exam").val(); 
    var ofc=$("#ofc").val(); 
    var ctrl =$("#ctrl").val(); 
   // var ins1 = ins.replace(/(?:\r\n|\r|\n)/g, '<br>');
    
    
    
             $(`#ins_${row}`).html(ins.replace(/(?:\r\n|\r|\n)/g, '<br>'));
    $(`#exam_${row}`).html(exm.replace(/(?:\r\n|\r|\n)/g, '<br>'));
    $(`#ofc_${row}`).html(ofc.replace(/(?:\r\n|\r|\n)/g, '<br>'));
    $(`#ctrl_${row}`).html(ctrl.replace(/(?:\r\n|\r|\n)/g, '<br>'));
            }
}
  function readFile() {
  
  if (!this.files || !this.files[0]) return;
    
  const FR = new FileReader();
    
  var rowCount = $('.page').length;
    
  FR.addEventListener("load", function(evt) {
    for (var row = 0; row < rowCount; row++) {  
      document.querySelector(`#sign_${row}`).src         = evt.target.result;
    //  document.querySelector(`#llogo_${row}`).src         = evt.target.result;
    //  document.querySelector(`#rlogo_${row}`).src         = evt.target.result;
    }
    document.querySelector("#b64").textContent = evt.target.result;
  }); 
    
  FR.readAsDataURL(this.files[0]);
  
}
  function readFile1() {
  
  if (!this.files || !this.files[0]) return;
    
  const FR = new FileReader();
   
  var rowCount = $('.page').length; 
    
  FR.addEventListener("load", function(evt) {
    for (var row = 0; row < rowCount; row++) {  
      //document.querySelector(`#sign_${row}`).src         = evt.target.result;
      document.querySelector(`#llogo_${row}`).src         = evt.target.result;
      //document.querySelector(`#rlogo_${row}`).src         = evt.target.result;
    }
    document.querySelector("#b64").textContent = evt.target.result;
  }); 
    
  FR.readAsDataURL(this.files[0]);
  
}
  function readFile2() {
  
  if (!this.files || !this.files[0]) return;
    
  const FR = new FileReader();
    
  var rowCount = $('.page').length;
    
  FR.addEventListener("load", function(evt) {
    for (var row = 0; row < rowCount; row++) {  
     // document.querySelector(`#sign_${row}`).src         = evt.target.result;
    //  document.querySelector(`#llogo_${row}`).src         = evt.target.result;
      document.querySelector(`#rlogo_${row}`).src         = evt.target.result;
    }
    document.querySelector("#b64").textContent = evt.target.result;
  }); 
    
  FR.readAsDataURL(this.files[0]);
  
}

document.querySelector("#inp").addEventListener("change", readFile);
document.querySelector("#llogo").addEventListener("change", readFile1);
document.querySelector("#rlogo").addEventListener("change", readFile2);

  
  
 