<script setup lang="ts">
import { read, utils, writeFileXLSX } from "xlsx";
import { ref } from "vue";

const column = ref<string>("Bin 1");
const headers = ref<Array<string>>([]);
const fileRef = ref<HTMLInputElement>();
const firstSheetName = ref<string>(`Sheet1`);
const sortedRows = ref<Array<Record<string, any>>>([]);

const handleDownload = () => {
  const workBook = utils.book_new();
  const workSheet = utils.json_to_sheet(sortedRows.value);
  utils.book_append_sheet(workBook, workSheet, firstSheetName.value);
  writeFileXLSX(workBook, `Sorted ${new Date().toDateString()}.xlsx`);
};

const handleFileSelect = async ({ target }: Event) => {
  const files: FileList = (target as any)?.files;
  const file = files.item(0);
  if (!file) return;
  const workBook = read(await file.arrayBuffer());

  firstSheetName.value = workBook.SheetNames[0]; // Reading only first sheet for simplicity
  const workSheet = workBook.Sheets[firstSheetName.value];
  const rows = utils.sheet_to_json<Record<string, string>>(workSheet);

  sortedRows.value = Object.entries<Array<Record<string, any>>>(
    rows
      .filter((x) => x && x[column.value] && (x[column.value].startsWith("s") || x[column.value].startsWith("S")))
      .reduce<Record<string, any>>((p, c: Record<string, string>, _, __, [k] = /[^\d]+/g.exec(c[column.value]) || ["NULL"]) => ((p[k] || (p[k] = [])).push(c), p), {})
  )
    .sort(([kp], [kc]) => (kp < kc ? -1 : kp > kc ? 1 : 0))
    .map(([_, v]) => {
      return v.sort((prev, next) => {
        // Creating unique header values
        Object.keys(next).forEach((key) => {
          if (!headers.value.includes(key)) {
            headers.value.push(key);
          }
        });
        const prevMatch = /\d+/g.exec(prev[column.value]);
        const nextMatch = /\d+/g.exec(next[column.value]);
        if (prevMatch?.length && nextMatch?.length) {
          if (+prevMatch[0] < +nextMatch[0]) return -1;
          if (+prevMatch[0] > +nextMatch[0]) return 1;
        }
        return 0;
      });
    })
    .flat<Array<Record<string, any>>>(1);
};
</script>

<template lang="pug">
form.form(@submit.prevent="handleDownload" ref="form")
  input.input(v-model.trim="column" placeholder="Column to sort by" required)
  button.button(@click="fileRef?.click" type="button") Upload
  button.button(type="submit" v-if="sortedRows.length") Download
  input(type="file" accept="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet" ref="fileRef" @change="handleFileSelect" hidden)
table.table(v-if="headers.length && sortedRows.length")
  thead
    tr.table-row
      th.table-header S.No.
      th.table-header(v-for="header of headers") {{ header }}
  tbody
    tr.table-row(v-for="(row, index) of sortedRows")
      td.table-data {{ index + 1 }}
      td.table-data(v-for="data of Object.values(row)") {{ data }}
</template>

<style lang="stylus" scoped>
.form
  display flex
  align-items center
  flex-direction column
  gap 1rem

  > *
    flex 1

  .button
    width 100%

  .input
    border 1px solid transparent
    transition border-color 0.25s
    border-radius 0.5rem
    padding 0.6rem

    &:hover
      border-color: #646cff;

  @media screen and (min-width: 768px)
    flex-direction row
</style>

<style lang="stylus" scoped>
.table
  border 2px solid #646cff
  border-radius 0.5rem
  margin-top 1rem
  width 100%

  &-header
    padding 0.25rem
    background #2a2a2a

  &-data
    text-align center

@media screen and (min-width: 768px)
  max-width 75%
</style>
