<template>
    <div @click="exportExcel">
		<slot>
		</slot>
	</div>
</template>

<script>
    import XLSX from 'xlsx/xlsx';

    export default {
        name: "vue-excel-xlsx",

        props: {
            columns: {
                type: Array,
                default: () => []
            },
            data: {
                type: Array,
                default: () => []
            },
            promisedData: {
                type: Promise
            },
            filename: {
                type: String,
                default: 'excel'
            },
            sheetname: {
                type: String,
                default: 'SheetName'
            }
		},

        methods: {
            async exportExcel() {
                let createXLSLFormatObj = [];
                let newXlsHeader = [];
                let vm = this;
                if (vm.columns.length === 0){
                    console.log("Add columns!");
                    return;
                }
                let data = vm.data;
                if (data.length === 0){
                    if (vm.promisedData){
                    	data = await vm.promisedData;
                    }
                    if (data.length === 0){
                    	console.log("Add data!");
                    	return;
                    }
                }

                vm.columns.forEach( function(value, index){
                	newXlsHeader.push(value.label);
                });

                createXLSLFormatObj.push(newXlsHeader);
                data.forEach(function(value, index){
                	let innerRowData = [];
                    vm.columns.forEach(function(val, index){
                    	if (val.dataFormat && typeof val.dataFormat === 'function') {
                            innerRowData.push(val.dataFormat(value[val.field]));
                        }else {
                            innerRowData.push(value[val.field]);
                        }
                    });
                    createXLSLFormatObj.push(innerRowData);
                });

                let filename = vm.filename + ".xlsx";

                let ws_name = vm.sheetname;

                let wb = XLSX.utils.book_new(),
                    ws = XLSX.utils.aoa_to_sheet(createXLSLFormatObj);
                XLSX.utils.book_append_sheet(wb, ws, ws_name);
                XLSX.writeFile(wb, filename);
            }
        }
    }
</script>
