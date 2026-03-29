import { NextRequest, NextResponse } from "next/server";
import { processFiles, MAX_FILE_SIZE } from "@/lib/excel-processor";

/** Sanitize dateLabel to prevent Excel formula injection (=, +, -, @, tab, CR). */
function sanitizeDateLabel(label: string): string {
  return label.replace(/^[=+\-@\t\r]+/, "");
}

export async function POST(request: NextRequest) {
  try {
    const formData = await request.formData();
    const inputRaw = formData.get("input");
    const outputRaw = formData.get("output");
    const dateLabelRaw = formData.get("dateLabel");
    const dateLabel =
      typeof dateLabelRaw === "string" ? sanitizeDateLabel(dateLabelRaw.trim()) : "";

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

    if (inputRaw.size === 0 || outputRaw.size === 0) {
      return NextResponse.json(
        { error: "Uploaded files must not be empty" },
        { status: 400 }
      );
    }

    if (inputRaw.size > MAX_FILE_SIZE || outputRaw.size > MAX_FILE_SIZE) {
      return NextResponse.json(
        {
          error: `File size must not exceed ${MAX_FILE_SIZE / 1024 / 1024} MB`,
        },
        { status: 413 }
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
      { error: "Failed to process files. Please check that both files are valid .xlsx files." },
      { status: 500 }
    );
  }
}
