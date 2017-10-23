'use strict'

//Includes Modules
const EventEmitter = require('events').EventEmitter;
const night = require('node-nightmare');
const XLSX = require('xlsx');
const utf8 = require('utf8');

//Config Values
const config = require("./config/config");

function callback(){
    console.log('hola');
}

//Web Process Class
class nightmare extends EventEmitter{
	constructor(){
		super();
		this.status = "start";
        var that = this;
	}
	webInfo(r){
		this.emit('wrote');
	}
	webEnd(){
		this.status = "end";
	}
	webProcess(d){
		night({
            openDevTools://{mode: 'detach'}, //Testing.
            show: true}) 
    		.goto(config.url)
    		.wait('body')
    		.show()
    		.type('#user_name', config.user)
		    .type('#user_password', config.pass)
		    .click('#login_button')
		    .wait(1000)
		    .wait('#search_form_clear')
		    .click('#search_form_clear')
		    .type('#numeroidentificacion_c_basic', d.Dni)
		    .click('#search_form_submit')
		    .wait(1000)
		    .wait('body')
		    .click('#button_select_all_top')
		    .wait(1000)
		    .wait('body')
		    .click('#mergeduplicates_listview_top')
		    .wait(3000)
		    .wait('#cancel_merge_button')
		    .click('#cancel_merge_button')
		    .wait(1000)
		    .wait('body')
            .emitEnd()
            .evaluate(function(){
                console.log('logg');
            })
            .end()
		    .then(function(ret) {
                console.log('log ::' + ret);
                callback();
		    });
            callback();
        //that.emit('register_complete', d);
	}
}
//

//GESTIO EXCEL
////////////////////////////////////////////////////////////////////////
class excelWeb extends EventEmitter {
  	constructor(file){
    	super();
    	this.indicadors = [ 'A' ];
    	this.file = file;
    	this.status = "start";
	
		this.columnes = [
			{ column: 'A', valor: 'Dni'},
		]
  	}
  	read(){
    	var workbook = XLSX.readFile(this.file);
    	var first_sheet_name = workbook.SheetNames[0];
    	var worksheet = workbook.Sheets[first_sheet_name];
    	var fila = 1;
    	var actualCell;
    	var that = this;
    	for (actualCell = worksheet['A'+fila]; actualCell; ){
      		var dades = [];
			for(let i=0; i < this.columnes.length; i++){
				let dada = this.columnes[i].valor;
				let col = this.columnes[i].column;
				col = col + fila;
				if (worksheet[col])
					dades[dada] = worksheet[col].v;
				else
					dades[dada] = "";
			}
      		that.emit('registre', dades);
      		fila++;
      		actualCell = worksheet['A'+fila];
    	}
    	that.status = "end";
		that.emit('end');
  	}
}

class doMigrarExcel {
	constructor() {
		var that = this;
		this.Web = new excelWeb(config.excelWeb);
		this.WebAuto = new nightmare();
		this.count_register = 0;
		
		this.WebAuto.on('wrote', function(d){
			that.popRegister();
		});
		this.Web.on('registre', function(r){
				that.pushRegister();
				that.WebAuto.webProcess(r);
			});
		this.WebAuto.on('register_complete', function(r){
				that.WebAuto.webInfo(r);
			});
		this.Web.on('end', function(){
			that.checkEnd()
		});
  	}
  	start(){
    	this.Web.read();
  	}
  	pushRegister(){
    	this.count_register++;
  	}
  	popRegister(){
    	this.count_register--;
    	this.checkEnd();
  	}
  	checkEnd() {
		if (this.count_register == 0 && this.Web.status == 'end'){
      		this.WebAuto.webEnd();
		}
  	}
}
////////////////////////////////////////////////////////////////////////



let myProgram = new doMigrarExcel();
myProgram.start();
