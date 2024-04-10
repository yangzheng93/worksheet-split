<script setup>
import { ref } from "vue";
import { groupBy } from "lodash-es";
import { read, utils, writeFile } from "xlsx";

const fileList = ref([]);
const toGenerate = async () => {
  if (fileList.value.length > 0) {
    const file = fileList.value[0];
    const ab = await file.arrayBuffer();
    // 返回一个workbook对象
    const wb = read(ab);

    // 把所有sheet的数据合并
    const data = wb.SheetNames.reduce((a, curSheetName) => {
      const curws = wb.Sheets[curSheetName];
      return a.concat(
        utils.sheet_to_json(curws, { defval: "", header: 0 }).map((i) => ({
          ...i,
          文件标题: i["文件标题"] || "Unknown",
          sheetName: curSheetName,
        })),
      );
    }, []);

    // 所有数据按照文件名称分组
    const nameGroup = groupBy(data, "文件标题");

    Object.keys(nameGroup).forEach((name) => {
      const list = nameGroup[name];
      const innerGroup = groupBy(list, "sheetName");

      const innerWorkBook = utils.book_new();
      Object.keys(innerGroup).forEach((sn) => {
        const worksheet = utils.json_to_sheet(
          innerGroup[sn].map((i) => {
            delete i.sheetName;
            return i;
          }),
        );
        utils.book_append_sheet(innerWorkBook, worksheet, sn);
      });

      writeFile(innerWorkBook, `${name}.xlsx`, { compression: true });
    });
  }
};

const beforeUpload = (file) => {
  fileList.value = [...(fileList.value || []), file];
  return false;
};
</script>

<template>
  <a-card :style="{ width: '100%' }">
    <a-space :size="20" direction="vertical">
      <a-upload
        :file-list="fileList"
        :max-count="1"
        :before-upload="beforeUpload"
      >
        <a-button> 选择文件 </a-button>
      </a-upload>

      <a-button
        type="primary"
        :disabled="fileList.length === 0"
        @click="toGenerate"
      >
        点击转换
      </a-button>
    </a-space>
  </a-card>
</template>

<style></style>
