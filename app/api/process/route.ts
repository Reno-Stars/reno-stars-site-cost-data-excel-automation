import { NextRequest, NextResponse } from "next/server";
import { processFiles, MAX_FILE_SIZE } from "@/lib/excel-processor";

export async function POST(request: NextRequest) {
  try {
    const formData = await request.formData();
    const inputRaw = formData.get("input");
    const outputRaw = formData.get("output");
    const dateLabelRaw = formData.get("dateLabel");
    const dateLabel =
      typeof dateLabelRaw === "string" ? dateLabelRaw.trim() : "";

    if (!(inputRaw instanceof File) || !(outputRaw instanceof File)) {
      return NextResponse.json(
        { error: "Both input and output files are required" },
        { status: 400 }
      );
    }

    if (!inputRaw.name.endsWith(".xlsx") || !outputRaw.name.endsWith(".xlsx")) {
      return NextResponse.json(
        { error: "Only .xlsx files are accepted" },
        { status: 400 }
      );
    }

    if (inputRaw.size > MAX_FILE_SIZE || outputRaw.size > MAX_FILE_SIZE) {
      return NextResponse.json(
        {
          error: `File size must not exceed ${MAX_FILE_SIZE / 1024 / 1024} MB`,
        },
        { status: 400 }
      );
    }

    if (!dateLabel) {
      return NextResponse.json(
        { error: "Date label (日期) is required" },
        { status: 400 }
      );
    }

    const inputBuffer = await inputRaw.arrayBuffer();
    const outputBuffer = await outputRaw.arrayBuffer();

    const { result, outputFile: processedBuffer } = await processFiles(
      inputBuffer,
      outputBuffer,
      dateLabel
    );

    return NextResponse.json({
      ...result,
      file: processedBuffer.toString("base64"),
      fileName: outputRaw.name.replace(".xlsx", "-updated.xlsx"),
    });
  } catch (error) {
    console.error("Processing error:", error);
    return NextResponse.json(
      {
        error:
          error instanceof Error ? error.message : "Failed to process files",
      },
      { status: 500 }
    );
  }
}
