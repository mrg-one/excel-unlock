/* eslint-disable @typescript-eslint/no-explicit-any */
/* eslint-disable @typescript-eslint/no-unsafe-argument */
/* eslint-disable @typescript-eslint/no-unsafe-call */
/* eslint-disable @typescript-eslint/no-unsafe-member-access */
/* eslint-disable @typescript-eslint/no-unsafe-assignment */
import Head from "next/head";

import JSZip from "jszip";
import { parseStringPromise, Builder } from "xml2js";
import React from "react";

export default function Home() {

  const handleFileUpload = async (
    event: React.ChangeEvent<HTMLInputElement>,
  ) => {
    const file = event.target.files?.[0];
    if (!file) return;

    try {
      const arrayBuffer = await file.arrayBuffer();
      const zip = await JSZip.loadAsync(arrayBuffer);

      const allSheets = zip.folder("xl/worksheets");

      if (allSheets === null) {
        alert("no sheets in workbook");
        return;
      }

      const sheetnames = Object.keys(allSheets.files).filter(
        (sh) => sh.startsWith("xl/worksheets/") && sh.endsWith(".xml"),
      )



      for (const sheetname of sheetnames) {
        console.log(sheetname);
        const sheetFile = zip.file(sheetname);
        if (!sheetFile) {
          alert(`${sheetname} not found!`);
          return;
        }

        const xmlContent = await sheetFile.async("string");
        const parsed = (await parseStringPromise(xmlContent)) as unknown as any;

        // Remove sheetProtection if it exists
        if (parsed.worksheet?.sheetProtection) {
          delete parsed.worksheet.sheetProtection;
        }

        // Build updated XML
        const builder = new Builder();
        const modifiedXml = builder.buildObject(parsed);

        // Replace file in archive
        zip.file(sheetname, modifiedXml);
      }

      // Repackage as Blob and trigger download
      const newBlob = await zip.generateAsync({ type: "blob" });
      const url = URL.createObjectURL(newBlob);
      const a = document.createElement("a");
      a.href = url;
      a.download = file.name.split(".xlsx")[0]!.concat("__unlocked.xlsx");
      a.click();
      URL.revokeObjectURL(url);
      (document.getElementById("fileuploader") as HTMLInputElement).value  = ""
    } catch (error) {
      console.error("Error unlocking sheet:", error);
      alert("Something went wrong while processing the file.");
    }
  };

  return (
    <>
      <Head>
        <title>Make-Excel-Write</title>
        <link rel="icon" href="/favicon.ico" />
      </Head>
      <main className="flex min-h-screen flex-col items-center justify-center bg-gradient-to-br from-black to-slate-900">
        <input
          id="fileuploader"
          type="file"
          accept=".xlsx"
          onChange={handleFileUpload}
          className={`rounded-4xl border-2 bg-white p-4 mb-4`}
        />
      </main>
    </>
  );
}
