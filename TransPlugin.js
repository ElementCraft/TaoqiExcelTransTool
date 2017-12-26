const xlsx = require('node-xlsx');
const fs = require("fs");

//==========================
// 声明和定义部分
//==========================

// 全局变量
var arrAllCars = [];
var arrResultCars = [];
var arrAllArea = [];

function CarNumber()
{
	this.wait = 0; // 待售
	this.have = 0; // 在库
	this.run = 0;  // 在途
}

function Car()
{
	this.brand = "";
	this.name = "";
	this.size = "";
	this.money = 0.0;
	this.color = "";
	this.state = "";
	this.nums = new Map();
	this.results = new Map();
}

//==========================
// 逻辑部分
//==========================

// 解析excel
const origin = xlsx.parse("origin.xlsx");
const data = origin[0].data;


var currentState = "待售";	// 记录当前读取到的车辆状态分类
var currentBrand = "";		// 记录当前读取到的车辆品牌
var currentName = "";		// 记录当前读取到的车辆品名

// 读取所有地区（最后有用
for (let j = 6; j < data[0].length; j++) {
	let area = data[0][j] ? data[0][j].trim() : undefined;

	if(area)
		arrAllArea.push(area);
}

// 读取车辆信息
for (let i = 2; i < data.length; i++) {
	const o = data[i];
	
	// 状态切换
	let state = o[0] ? o[0].trim() : undefined;
	if(state && state == "在库处理车辆") 
		currentState = "在库";
	else if(state && state == "在途车辆") 
		currentState = "在途";

	// 品牌
	let brand = o[1] ? o[1].trim() : undefined;
	if(brand)
		currentBrand = brand;

	// 品名
	let name = o[2] ? o[2].trim() : undefined;
	if(name)
		currentName = name;

	var car = new Car();
	car.state = currentState;
	car.brand = currentBrand;
	car.name = currentName;
	car.size = o[3] ? o[3].trim() : "";
	car.money = o[4];
	car.color = o[5] ? o[5].trim() : "";

	// 删除款式末尾颜色
	if(car.size.match(/^.+、.$/g)){
		car.size = car.size.replace(/、.$/, "");
	}
	
	// 获取所有地区车辆数量
	for (let j = 6; j < data[0].length; j++) {
		let area = data[0][j] ? data[0][j].trim() : undefined;

		if(area){
			car.nums.set(area, o[j] ? o[j] : 0);
		}
	}

	arrAllCars.push(car);
}

// 重新汇总结果
arrAllCars.forEach(obj => {
	var bo = true;

	// 遍历新数组 寻找相同车型
	arrResultCars.forEach(result => {
		// 相同时合并
		if(obj.brand == result.brand 
			&& obj.money == result.money
			&& obj.color == result.color
			&& obj.name == result.name
			&& obj.size == result.size
		){
			// 遍历地区map 合并
			obj.nums.forEach(function (item, key, mapObj) {
				var temp = result.results.get(key);
				if(obj.state == "在库") temp.have = item;
				if(obj.state == "在途") temp.run = item;

				result.results.set(temp);
			});
			bo = false;
		}
	});

	// 未找到相同元素 初始化results这个map后 push到数组中
	if(bo)
	{
		arrAllArea.forEach(item => {
			let oo = new CarNumber();
			let ii = obj.nums.get(item);

			if(obj.state == "待售") 
				oo.wait = ii;
			else if(obj.state == "在库") 
				oo.have = ii;
			else if(obj.state == "在途") 
				oo.run = ii;
			
			obj.results.set(item, oo);
		});

		arrResultCars.push(obj);
	}
});

//========================================================
// 存入文件中
//========================================================

let final = []
let L1 = [undefined,undefined,undefined,undefined,undefined];
let L2 = ["品牌","品名","规格","指导价","颜色"]
let range = []

// 第一、二行
arrAllArea.forEach( (item , i) => {
	L1.push(item);
	L1.push(undefined);
	L1.push(undefined);
	L1.push(undefined);
	L1.push(undefined);

	L2.push("小计");
	L2.push("待售");
	L2.push("在库");
	L2.push("在途");
	L2.push("待采");

	range.push({s: {c: 5 * i, r:0 }, e: {c: 5 * i + 4, r:0}});
});

//构造最终数组
final.push(L1);
final.push(L2);

arrResultCars.forEach((item, i) => {
	var arr = []
	arr.push(item.brand);
	arr.push(item.name);
	arr.push(item.size);
	arr.push(item.money);
	arr.push(item.color);

	arrAllArea.forEach( (area) => {
		var oo = item.results.get(area);
		var num = oo.wait + oo.have + oo.run;

		arr.push(num == 0 ? "-" : num);
		arr.push(oo.wait == 0 ? undefined : oo.wait);
		arr.push(oo.have == 0 ? undefined : oo.have);
		arr.push(oo.run == 0 ? undefined : oo.run);
		arr.push(undefined);
	});

	final.push(arr);
});

let sheet = xlsx.build([{
	name: 'sheet1',
	data: final
}], {'!merges': range });

fs.writeFile('./result/result.xlsx', sheet, function(err) {  
	if (err)  
		console.log(err);
});