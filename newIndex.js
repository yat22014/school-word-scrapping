const axios = require('axios');
const cheerio = require('cheerio');
const fs = require('fs');
const ExcelJS = require('exceljs');

const targetUrls = [
  'https://www.edb.gov.hk/en/student-parents/sch-info/sch-search/schlist-by-district/school-list-cw.html',
  'https://www.edb.gov.hk/en/student-parents/sch-info/sch-search/schlist-by-district/school-list-hke.html',
  'https://www.edb.gov.hk/en/student-parents/sch-info/sch-search/schlist-by-district/school-list-i.html',
  'https://www.edb.gov.hk/en/student-parents/sch-info/sch-search/schlist-by-district/school-list-sou.html',
  'https://www.edb.gov.hk/en/student-parents/sch-info/sch-search/schlist-by-district/school-list-wch.html',
  'https://www.edb.gov.hk/en/student-parents/sch-info/sch-search/schlist-by-district/school-list-kc.html',
  'https://www.edb.gov.hk/en/student-parents/sch-info/sch-search/schlist-by-district/school-list-kt.html',
  'https://www.edb.gov.hk/en/student-parents/sch-info/sch-search/schlist-by-district/school-list-sk.html',
  'https://www.edb.gov.hk/en/student-parents/sch-info/sch-search/schlist-by-district/school-list-ssp.html',
  'https://www.edb.gov.hk/en/student-parents/sch-info/sch-search/schlist-by-district/school-list-wts.html',
  'https://www.edb.gov.hk/en/student-parents/sch-info/sch-search/schlist-by-district/school-list-ytm.html',
  'https://www.edb.gov.hk/en/student-parents/sch-info/sch-search/schlist-by-district/school-list-n.html',
  'https://www.edb.gov.hk/en/student-parents/sch-info/sch-search/schlist-by-district/school-list-st.html',
  'https://www.edb.gov.hk/en/student-parents/sch-info/sch-search/schlist-by-district/school-list-tp.html',
  'https://www.edb.gov.hk/en/student-parents/sch-info/sch-search/schlist-by-district/school-list-kwt.html',
  'https://www.edb.gov.hk/en/student-parents/sch-info/sch-search/schlist-by-district/school-list-tw.html',
  'https://www.edb.gov.hk/en/student-parents/sch-info/sch-search/schlist-by-district/school-list-tm.html',
  'https://www.edb.gov.hk/en/student-parents/sch-info/sch-search/schlist-by-district/school-list-yl.html',
]
let allData = [];
let targetPromises = [];
targetUrls.forEach((targetUrl) => {
  targetPromises.push(new Promise((r1, r2) => {
    axios.get(targetUrl).then((response) => {
      getData(response.data);
      r1(true);
    }).catch((error) => {
      console.log(9, error);
    });

  }));
});

Promise.all(targetPromises).then(() => {
  createExcel();
});

let getData = html => {
  data = [];
  const $ = cheerio.load(html);
  // $('body>table>tbody>tr:nth-child(2) td').each((i, elem) => {
  //   console.log(18, i, $(elem).text().trim());
  //   // data.push({
  //   //   title: $(elem).text(),
  //   //   link: $(elem).find('a.storylink').attr('href')
  //   // });
  // });
  $('body > table').each((i, elem) => {
    // console.log(25, i, elem.children, $(elem));
    // console.log(25, i, $(elem).children('tbody').children('tr').eq(2).children('td').eq(1).text());
    let districtText = $(elem).children('tbody').children('tr').eq(1).children('td').eq(0).text().trim();
    schoolDistrict = districtText.replace(/ +/g, ' ');
    let schoolTypeText = $(elem).children('tbody').children('tr').eq(1).children('td').eq(1).text().trim();
    schoolType = schoolTypeText.replace(/ +/g, ' ');
    let schoolTable = $(elem).children('tbody').children('tr').eq(2).children('td').eq(0).children('table').eq(0).children('tbody').eq(0).children('tr');
    schoolTable.each((index, schoolTableTr) => {
      if (index !== 0) {
        //push data here
        let thisRowSpan = $(schoolTableTr).children('td').eq(0).prop("rowspan");
        if (thisRowSpan == 2 || thisRowSpan == 1) {
          let schoolEnglishName = $(schoolTableTr).children('td').eq(1).children('table').eq(0).children('tbody').eq(0).children('tr').eq(0).children('td').eq(0).text();
          let schoolEnglishAddress = $(schoolTableTr).children('td').eq(1).children('table').eq(0).children('tbody').eq(0).children('tr').eq(1).children('td').eq(1).text().trim();
          let schoolChineseName = $(schoolTableTr).children('td').eq(1).children('table').eq(0).children('tbody').eq(0).children('tr').eq(2).children('td').eq(0).text();
          let schoolChineseAddress = $(schoolTableTr).children('td').eq(1).children('table').eq(0).children('tbody').eq(0).children('tr').eq(3).children('td').eq(1).text();
          // console.log(41, schoolEnglishName, schoolEnglishAddress, schoolChineseName, schoolChineseAddress);
          let schoolTel = $(schoolTableTr).children('td').eq(2).children('table').eq(0).children('tbody').eq(0).children('tr').eq(0).children('td').eq(0).text();
          let schoolFax = $(schoolTableTr).children('td').eq(2).children('table').eq(0).children('tbody').eq(0).children('tr').eq(1).children('td').eq(0).text();

          // console.log(42, $(schoolTableTr).children('td').eq(3).children('table').eq(0).children('tbody').eq(0).children('tr').eq(0).children('td').eq(0).html());
          // console.log(42, $(schoolTableTr).children('td').eq(3).children('table').eq(0).children('tbody').eq(0).children('tr').eq(0).children('td').eq(0).text().trim());
          let schoolSupervisorText = $(schoolTableTr).children('td').eq(3).children('table').eq(0).children('tbody').eq(0).children('tr').eq(0).children('td').eq(0).text().trim();
          let schoolSupervisor = schoolSupervisorText.replace(/SMC/g, 'SMC ');
          schoolSupervisor = schoolSupervisor.replace(/\t+/g, ' ');

          let schoolPrincipalText = $(schoolTableTr).children('td').eq(3).children('table').eq(0).children('tbody').eq(0).children('tr').eq(1).children('td').eq(0).text().trim();
          schoolPrincipal = schoolPrincipalText.replace(/\t+/g, ' ');

          let schoolSexTypeText = $(schoolTableTr).children('td').eq(4).text().trim();
          schoolSexType = schoolSexTypeText.replace(/\t+|\*+/g, '');

          data.push({
            schoolDistrict: schoolDistrict,
            schoolType: schoolType,
            schoolEnglishName: schoolEnglishName,
            schoolEnglishAddress: schoolEnglishAddress,
            schoolChineseName: schoolChineseName,
            schoolChineseAddress: schoolChineseAddress,
            schoolTel: schoolTel,
            schoolFax: schoolFax,
            schoolSupervisor: schoolSupervisor,
            schoolPrincipal: schoolPrincipal,
            schoolSexType: schoolSexType
          });

        }
        // console.log(35, $(schoolTableTr).children('td').eq(1).children('table').eq(0).children('tbody').eq(0).children('tr').eq(0).children('td').eq(0).text());
      }
    });
  });
  // console.log(22, data);
  // let jsonString = JSON.stringify(data);

  allData.push(data);
}

function createExcel() {

    let workbook = new ExcelJS.Workbook();
    workbook.creator = 'Frank';
    // workbook.lastModifiedBy = 'Her';
    workbook.created = new Date(2020, 4, 28);
    workbook.modified = new Date();
    // workbook.lastPrinted = new Date(2016, 9, 27);
    workbook.views = [
      {
        x: 0, y: 0, width: 10000, height: 20000,
        firstSheet: 0, activeTab: 1, visibility: 'visible'
      }
    ];


    allData.forEach((sheetData, index) => {
      var schoolSheet = workbook.addWorksheet(sheetData[0].schoolDistrict, {properties:{tabColor:{argb:'FFC0000'}}});
      schoolSheet.columns = [
        { header: 'District', key: 'district', width: 30 },
        { header: 'Type', key: 'type', width: 30 },
        { header: 'English name', key: 'englishName', width: 30 },
        { header: 'English address', key: 'englishAddress', width: 30 },
        { header: 'Chinese name', key: 'chineseName', width: 30 },
        { header: 'Chinese address', key: 'chineseAddress', width: 30 },
        { header: 'Tel', key: 'tel', width: 30 },
        { header: 'Fax', key: 'fax', width: 30 },
        { header: 'Supervisor', key: 'supervisor', width: 30 },
        { header: 'Principal', key: 'principal', width: 30 },
        { header: 'Sex type', key: 'sexType', width: 30 },
      ];

      sheetData.forEach((eachData, index) => {
        schoolSheet.addRow({
          district: eachData.schoolDistrict,
          type: eachData.schoolType,
          englishName: eachData.schoolEnglishName,
          englishAddress: eachData.schoolEnglishAddress,
          chineseName: eachData.schoolChineseName,
          chineseAddress: eachData.schoolChineseAddress,
          tel: eachData.schoolTel,
          fax: eachData.schoolFax,
          supervisor: eachData.schoolSupervisor,
          principal: eachData.schoolPrincipal,
          sexType: eachData.schoolSexType,
        });
      });

    });

    workbook.xlsx.writeFile('school'+ new Date().getTime() + '.xlsx').then(() => {console.log(283);});
}
