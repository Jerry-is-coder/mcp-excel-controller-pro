#!/usr/bin/env node
const { McpServer } = require("@modelcontextprotocol/sdk/server/mcp.js");
const {
  StdioServerTransport,
} = require("@modelcontextprotocol/sdk/server/stdio.js");
const { z } = require("zod");
const ExcelJS = require("exceljs");
const fs = require("fs");
const path = require("path");
const { exec, spawn } = require("child_process");

// 새로운 MCP 서버 생성
const server = new McpServer({
  name: "ExcelControllerPro",
  version: "1.0.0",
});

// 명령 실행 Promise 래핑
const execPromise = (cmd) => {
  return new Promise((resolve, reject) => {
    exec(cmd, (error, stdout, stderr) => {
      if (error) {
        reject(error);
        return;
      }
      resolve(stdout);
    });
  });
};

// Excel 앱 실행 함수
async function openExcelFile(filePath) {
  try {
    // 파일 경로 정규화
    const normalizedPath = path.resolve(filePath);

    // 엑셀 앱 시작
    const process = spawn("start", ["excel", `"${normalizedPath}"`], {
      shell: true,
      detached: true,
      stdio: "ignore",
    });

    // 프로세스를 추적하지 않음
    process.unref();

    return true;
  } catch (error) {
    return false;
  }
}

// 파일이 열려있는지 확인하는 함수
function isFileOpen(filePath) {
  try {
    const fd = fs.openSync(filePath, "r+");
    fs.closeSync(fd);
    return false;
  } catch (e) {
    return true;
  }
}

// 임시 PowerShell 스크립트 실행 함수
async function runPowerShellScript(scriptContent) {
  const timestamp = Date.now();
  const scriptPath = path.join(process.cwd(), `excel_script_${timestamp}.ps1`);

  try {
    // 스크립트를 파일로 저장
    fs.writeFileSync(scriptPath, scriptContent);

    // PowerShell 스크립트 실행
    const result = await execPromise(
      `powershell -ExecutionPolicy Bypass -NoLogo -NonInteractive -File "${scriptPath}" 2> NUL`
    );

    return result;
  } catch (error) {
    throw error;
  } finally {
    // 임시 파일 삭제
    try {
      fs.unlinkSync(scriptPath);
    } catch (err) {
      // 파일 삭제 오류는 무시
    }
  }
}

// Excel 셀 참조(예: "A1", "B5")를 행과 열 번호로 변환하는 함수
function cellReferenceToRowCol(cellReference) {
  // 알파벳과 숫자 부분 분리 (예: "A11" -> "A"와 "11")
  const match = cellReference.match(/([A-Za-z]+)([0-9]+)/);
  if (!match) {
    throw new Error(`유효하지 않은 셀 참조: ${cellReference}`);
  }

  const colStr = match[1].toUpperCase();
  const row = parseInt(match[2], 10);

  // 알파벳을 열 번호로 변환 (A=1, B=2, ..., Z=26, AA=27, ...)
  let col = 0;
  for (let i = 0; i < colStr.length; i++) {
    col = col * 26 + (colStr.charCodeAt(i) - 64);
  }

  return { row, col };
}

// Excel 프로세스 확인 및 대기 함수
async function waitForExcelAvailable(maxAttempts = 3) {
  for (let i = 0; i < maxAttempts; i++) {
    try {
      const result = await execPromise(
        'powershell -Command "Get-Process excel -ErrorAction SilentlyContinue | Out-String"'
      );
      if (result && result.trim()) {
        // Excel이 실행 중이면 잠시 대기
        await new Promise((resolve) => setTimeout(resolve, 1000));
        return true;
      }
    } catch (e) {
      // 오류 무시
    }
    await new Promise((resolve) => setTimeout(resolve, 500));
  }
  return false;
}

// 열려있는 Excel 파일의 데이터 업데이트 함수 (재시도 로직 포함)
async function updateExcelWithRetry(filePath, sheetName, data, maxAttempts = 3) {
  for (let i = 0; i < maxAttempts; i++) {
    try {
      // Excel 프로세스 확인 및 대기
      await waitForExcelAvailable();
      
      // 정규화된 경로로 변환
      const normalizedPath = path.resolve(filePath);

      // 데이터 형식 검증
      if (!Array.isArray(data) || data.length === 0) {
        return {
          success: false,
          message: "유효한 데이터가 아닙니다. 2차원 배열 형식이어야 합니다.",
        };
      }

      // PowerShell 스크립트 작성 - 디버그 출력 완전 제거
      let psScript = `
        # 모든 오류 메시지 숨기기 (중요한 오류만 캡처)
        $ErrorActionPreference = "SilentlyContinue"
        
        try {
            # 기존 Excel 인스턴스 가져오기
            $excel = $null
            try {
                $excel = [Runtime.InteropServices.Marshal]::GetActiveObject("Excel.Application")
            } catch {
                # 기존 인스턴스가 없으면 새로 생성
                $excel = New-Object -ComObject Excel.Application
            }
            
            $excel.Visible = $true
            $excel.DisplayAlerts = $false
            
            # 이미 열려있는 워크북 찾기
            $workbook = $null
            foreach ($wb in $excel.Workbooks) {
                if ($wb.FullName -eq "${normalizedPath.replace(
                  /\\/g,
                  "\\\\"
                )}") {
                    $workbook = $wb
                    break
                }
            }
            
            # 워크북이 없으면 열기
            if ($workbook -eq $null) {
                $workbook = $excel.Workbooks.Open("${normalizedPath.replace(
                  /\\/g,
                  "\\\\"
                )}")
            }
            
            # 워크북 활성화 - 현재 활성 워크북으로 설정
            $workbook.Activate()
            
            # 시트 선택
            $worksheet = $null
        `;

      if (sheetName) {
        psScript += `
            # 지정된 시트 찾기 시도
            try {
                $worksheet = $workbook.Worksheets("${sheetName}")
            } catch {
                # 지정된 시트가 없으면 첫 번째 시트 사용
                $worksheet = $workbook.Worksheets.Item(1)
            }
            `;
      } else {
        psScript += `
            # 시트 이름을 지정하지 않았으므로 활성 시트 사용
            $worksheet = $workbook.ActiveSheet
            `;
      }

      psScript += `
            # 시트 활성화 - 현재 활성 시트로 설정
            $worksheet.Activate()
            
            # 데이터가 있는 사용된 범위를 모두 지우기
            # 주의: 기존 데이터가 없을 수도 있으므로 예외 처리
            try {
                $usedRange = $worksheet.UsedRange
                if ($usedRange -ne $null -and $usedRange.Cells.Count -gt 0)
            } catch {
                # 계속 진행
            }
            
            # 데이터를 배열로 설정
            $rowCount = ${data.length}
            $colCount = ${Math.max(...data.map((row) => row.length))}
            
            # 데이터를 한 번에 설정할 범위 생성
            $targetRange = $worksheet.Range($worksheet.Cells(1, 1), $worksheet.Cells($rowCount, $colCount))
            
            # 2차원 배열 생성
            $dataArray = New-Object 'object[,]' $rowCount, $colCount
        `;

      // 데이터 행과 열에 대한 설정
      for (let rowIndex = 0; rowIndex < data.length; rowIndex++) {
        const row = data[rowIndex];
        for (let colIndex = 0; colIndex < row.length; colIndex++) {
          const cellValue = row[colIndex];
          // 값 타입에 따른 처리
          if (typeof cellValue === "string") {
            // 문자열은 따옴표로 묶고 특수문자 처리
            const escapedValue = cellValue
              .replace(/'/g, "''")
              .replace(/"/g, '""');
            psScript += `
            $dataArray[${rowIndex}, ${colIndex}] = '${escapedValue}'`;
          } else if (cellValue === null || cellValue === undefined) {
            // null/undefined는 빈 문자열로
            psScript += `
            $dataArray[${rowIndex}, ${colIndex}] = ''`;
          } else if (typeof cellValue === "number") {
            // 숫자는 그대로
            psScript += `
            $dataArray[${rowIndex}, ${colIndex}] = ${cellValue}`;
          } else if (typeof cellValue === "boolean") {
            // 불리언 값 처리
            psScript += `
            $dataArray[${rowIndex}, ${colIndex}] = $${cellValue}`;
          } else {
            // 기타 값은 문자열로 변환
            psScript += `
            $dataArray[${rowIndex}, ${colIndex}] = '${String(cellValue).replace(
              /'/g,
              "''"
            )}'`;
          }
        }
      }

      psScript += `
            
            # 데이터 배열을 범위에 한 번에 설정
            $targetRange.Value2 = $dataArray
            
            # 변경 사항이 보이도록 셀 A1로 이동
            $worksheet.Cells(1, 1).Select()
            
            # 저장 확인
            $workbook.Save()
            
            # 자동 맞춤 적용 (열 너비 자동 조절)
            $usedRange = $worksheet.UsedRange
            $usedRange.Columns.AutoFit() | Out-Null
            
            # 성공 메시지 - 표준 출력으로만 한 줄 출력
            Write-Output "SUCCESS: Worksheet updated with ${data.length} rows of data and saved"
        } catch {
            # 오류 내용 - 표준 출력으로만 한 줄 출력
            Write-Output "ERROR: $($_.Exception.Message)"
        } finally {
            # COM 객체 참조 해제 (Excel 프로그램은 종료하지 않음)
            if ($worksheet -ne $null) {
                [System.Runtime.Interopservices.Marshal]::ReleaseComObject($worksheet) | Out-Null
            }
            if ($workbook -ne $null) {
                [System.Runtime.Interopservices.Marshal]::ReleaseComObject($workbook) | Out-Null
            }
            if ($excel -ne $null) {
                [System.Runtime.Interopservices.Marshal]::ReleaseComObject($excel) | Out-Null
            }
            [System.GC]::Collect()
            [System.GC]::WaitForPendingFinalizers()
        }
        `;

      // PowerShell 스크립트 실행
      const result = await runPowerShellScript(psScript);

      // 결과 확인
      if (result && result.includes("SUCCESS")) {
        return { success: true, message: result.trim() };
      } else {
        if (i === maxAttempts - 1) {
          // 마지막 시도에서 실패했을 때만 오류 반환
          return {
            success: false,
            message: result
              ? result.trim()
              : "스크립트 실행 중 오류가 발생했습니다.",
          };
        }
        // 아니면 재시도를 위해 계속 진행
      }
    } catch (error) {
      console.error(`시도 ${i + 1}/${maxAttempts} 실패:`, error.message);
      if (i === maxAttempts - 1) {
        // 마지막 시도에서 실패했을 때만 오류 반환
        return {
          success: false,
          message: `범위 데이터 업데이트 오류: ${error.message}`,
        };
      }
    }
    // 다음 시도 전 잠시 대기 (점진적으로 대기 시간 증가)
    await new Promise((resolve) => setTimeout(resolve, 1000 * (i + 1)));
  }
  
  // 모든 시도 실패 시 기본 오류 반환
  return {
    success: false,
    message: `${maxAttempts}번 시도 후에도 Excel 파일 업데이트 실패`,
  };
}

// 특정 셀들만 업데이트하는 함수
async function updateSpecificCells(filePath, sheetName, cellData, maxAttempts = 3) {
  for (let i = 0; i < maxAttempts; i++) {
    try {
      // Excel 프로세스 확인 및 대기
      await waitForExcelAvailable();
      
      // 정규화된 경로로 변환
      const normalizedPath = path.resolve(filePath);

      // PowerShell 스크립트 작성
      let psScript = `
        # 모든 오류 메시지 숨기기 (중요한 오류만 캡처)
        $ErrorActionPreference = "SilentlyContinue"
        
        try {
            # 기존 Excel 인스턴스 가져오기
            $excel = $null
            try {
                $excel = [Runtime.InteropServices.Marshal]::GetActiveObject("Excel.Application")
            } catch {
                # 기존 인스턴스가 없으면 새로 생성
                $excel = New-Object -ComObject Excel.Application
            }
            
            $excel.Visible = $true
            $excel.DisplayAlerts = $false
            
            # 이미 열려있는 워크북 찾기
            $workbook = $null
            foreach ($wb in $excel.Workbooks) {
                if ($wb.FullName -eq "${normalizedPath.replace(
                  /\\/g,
                  "\\\\"
                )}") {
                    $workbook = $wb
                    break
                }
            }
            
            # 워크북이 없으면 열기
            if ($workbook -eq $null) {
                $workbook = $excel.Workbooks.Open("${normalizedPath.replace(
                  /\\/g,
                  "\\\\"
                )}")
            }
            
            # 워크북 활성화 - 현재 활성 워크북으로 설정
            $workbook.Activate()
            
            # 시트 선택
            $worksheet = $null
        `;

      if (sheetName) {
        psScript += `
            # 지정된 시트 찾기 시도
            try {
                $worksheet = $workbook.Worksheets("${sheetName}")
            } catch {
                # 지정된 시트가 없으면 첫 번째 시트 사용
                $worksheet = $workbook.Worksheets.Item(1)
            }
            `;
      } else {
        psScript += `
            # 시트 이름을 지정하지 않았으므로 활성 시트 사용
            $worksheet = $workbook.ActiveSheet
            `;
      }

      psScript += `
            # 시트 활성화 - 현재 활성 시트로 설정
            $worksheet.Activate()
            
            # 각 셀 업데이트
            $updateCount = 0
            $errorCount = 0
        `;

      // 각 셀 업데이트 코드 추가
      for (const cell of cellData) {
        try {
          let cellReference = cell.cellReference;
          let value = cell.value;
          
          // 셀 참조를 행과 열 번호로 변환 시도
          let row, col;
          
          try {
            const result = cellReferenceToRowCol(cellReference);
            row = result.row;
            col = result.col;
          } catch (err) {
            // 셀 참조 변환 오류 처리
            psScript += `
            Write-Output "ERROR_CELL: 유효하지 않은 셀 참조 - ${cellReference}"
            $errorCount++
            `;
            continue;
          }
          
          // 값 타입에 따른 처리
          let valueStr;
          if (typeof value === "string") {
            valueStr = `"${value.replace(/"/g, '""')}"`;
          } else if (value === null || value === undefined) {
            valueStr = '""';
          } else if (typeof value === "boolean") {
            valueStr = value ? "$true" : "$false";
          } else {
            valueStr = value;
          }

          // 셀 업데이트 코드 추가
          psScript += `
          try {
            $worksheet.Cells(${row}, ${col}).Value2 = ${valueStr}
            $updateCount++
          } catch {
            $errorCount++
            Write-Output "ERROR_CELL: 셀 ${cellReference} 업데이트 중 오류 발생"
          }
          `;
        } catch (error) {
          // 코드 생성 중 오류는 무시하고 계속 진행
          continue;
        }
      }

      psScript += `
            # 저장 확인
            $workbook.Save()
            
            # 성공 메시지 - 표준 출력으로만 한 줄 출력
            Write-Output "SUCCESS: 총 $updateCount 개 셀이 업데이트됨 (오류: $errorCount 개)"
        } catch {
            # 오류 내용 - 표준 출력으로만 한 줄 출력
            Write-Output "ERROR: $($_.Exception.Message)"
        } finally {
            # COM 객체 참조 해제 (Excel 프로그램은 종료하지 않음)
            if ($worksheet -ne $null) {
                [System.Runtime.Interopservices.Marshal]::ReleaseComObject($worksheet) | Out-Null
            }
            if ($workbook -ne $null) {
                [System.Runtime.Interopservices.Marshal]::ReleaseComObject($workbook) | Out-Null
            }
            if ($excel -ne $null) {
                [System.Runtime.Interopservices.Marshal]::ReleaseComObject($excel) | Out-Null
            }
            [System.GC]::Collect()
            [System.GC]::WaitForPendingFinalizers()
        }
        `;

      // PowerShell 스크립트 실행
      const result = await runPowerShellScript(psScript);

      // 결과 확인
      if (result && result.includes("SUCCESS")) {
        return { success: true, message: result.trim() };
      } else {
        if (i === maxAttempts - 1) {
          // 마지막 시도에서 실패했을 때만 오류 반환
          return {
            success: false,
            message: result
              ? result.trim()
              : "스크립트 실행 중 오류가 발생했습니다.",
          };
        }
        // 아니면 재시도를 위해 계속 진행
      }
    } catch (error) {
      console.error(`시도 ${i + 1}/${maxAttempts} 실패:`, error.message);
      if (i === maxAttempts - 1) {
        // 마지막 시도에서 실패했을 때만 오류 반환
        return {
          success: false,
          message: `셀 데이터 업데이트 오류: ${error.message}`,
        };
      }
    }
    // 다음 시도 전 잠시 대기 (점진적으로 대기 시간 증가)
    await new Promise((resolve) => setTimeout(resolve, 1000 * (i + 1)));
  }
  
  // 모든 시도 실패 시 기본 오류 반환
  return {
    success: false,
    message: `${maxAttempts}번 시도 후에도 Excel 셀 업데이트 실패`,
  };
}

// 통합된 Excel 업데이트 도구
// 통합된 Excel 업데이트 도구
server.tool(
  "update_excel",
  "Excel 파일의 데이터를 업데이트합니다. 전체 범위 또는 특정 셀을 지정하여 업데이트할 수 있습니다.",
  {
    filePath: z.string().describe("엑셀 파일의 경로"),
    sheetName: z
      .string()
      .optional()
      .describe("시트 이름 (지정하지 않으면 활성 시트 또는 첫 번째 시트)"),
    // 세 가지 업데이트 모드 중 하나 선택
    mode: z
      .enum(["full", "append", "cells"])
      .default("full")
      .describe("업데이트 모드 (full: 덮어쓰기, append: 추가, cells: 특정 셀)"),
    data: z
      .array(z.array(z.any()))
      .optional()
      .describe("2차원 배열 형태의 데이터 (full 또는 append 모드에서 사용)"),
    cellData: z
      .array(
        z.object({
          cellReference: z.string(),
          value: z.any(),
        })
      )
      .optional()
      .describe("업데이트할 셀 데이터 배열 (cells 모드에서 사용)"),
    openFile: z
      .boolean()
      .optional()
      .default(false)
      .describe("작업 후 Excel로 파일을 열지 여부"),
    createBackup: z
      .boolean()
      .optional()
      .default(true)
      .describe("작업 전 백업 생성 여부"),
    smartMerge: z
      .boolean()
      .optional()
      .default(true)
      .describe("스마트 데이터 병합 모드 (데이터 구조 분석 및 적절한 병합, append 모드에서만 사용)"),
  },
  async ({
    filePath,
    sheetName,
    mode = "full",
    data,
    cellData,
    openFile = false,
    createBackup = true,
    smartMerge = true,
  }) => {
    try {
      // 파일 존재 확인
      if (!fs.existsSync(filePath)) {
        return {
          content: [
            { type: "text", text: `파일을 찾을 수 없습니다: ${filePath}` },
          ],
          isError: true,
        };
      }

      // 모드에 따른 데이터 유효성 검사
      if ((mode === "full" || mode === "append") && (!Array.isArray(data) || data.length === 0)) {
        // 데이터가 없지만 파일을 열기만 원하는 경우
        if (openFile) {
          const success = await openExcelFile(filePath);
          if (success) {
            return {
              content: [
                {
                  type: "text",
                  text: `엑셀 파일이 성공적으로 열렸습니다: ${filePath}`,
                },
              ],
            };
          } else {
            return {
              content: [
                {
                  type: "text",
                  text: `엑셀 파일을 열지 못했습니다: ${filePath}`,
                },
              ],
              isError: true,
            };
          }
        }

        return {
          content: [
            { type: "text", text: "유효한 데이터가 제공되지 않았습니다." },
          ],
          isError: true,
        };
      }

      if (mode === "cells" && (!Array.isArray(cellData) || cellData.length === 0)) {
        // 데이터가 없지만 파일을 열기만 원하는 경우
        if (openFile) {
          const success = await openExcelFile(filePath);
          if (success) {
            return {
              content: [
                {
                  type: "text",
                  text: `엑셀 파일이 성공적으로 열렸습니다: ${filePath}`,
                },
              ],
            };
          } else {
            return {
              content: [
                {
                  type: "text",
                  text: `엑셀 파일을 열지 못했습니다: ${filePath}`,
                },
              ],
              isError: true,
            };
          }
        }

        return {
          content: [
            { type: "text", text: "유효한 셀 데이터가 제공되지 않았습니다." },
          ],
          isError: true,
        };
      }

      // 백업 디렉토리 및 파일 생성 (작업 전)
      let backupPath = "";
      if (createBackup) {
        // 파일 이름과 확장자 분리
        const fileDir = path.dirname(filePath);
        const fileName = path.basename(filePath);
        const fileNameWithoutExt = path.basename(
          fileName,
          path.extname(fileName)
        );
        const fileExt = path.extname(fileName);

        // 백업 디렉토리 경로 생성
        const backupDir = path.join(fileDir, `log_${fileNameWithoutExt}`);

        // 백업 디렉토리가 없으면 생성
        if (!fs.existsSync(backupDir)) {
          fs.mkdirSync(backupDir, { recursive: true });
        }

        // 현지 시간대로 형식화
        const now = new Date();
        const options = {
          year: "numeric",
          month: "2-digit",
          day: "2-digit",
          hour: "2-digit",
          minute: "2-digit",
          second: "2-digit",
          hour12: false,
        };

        // 시스템의 현지 시간대를 사용
        const localTime = new Intl.DateTimeFormat(undefined, options).format(now);

        // 형식 정리 (yyyy-MM-dd HH:mm:ss → yyyyMMdd_HHmmss)
        const timestamp = localTime
          .replace(/[\/\-\,\s]/g, "") // 슬래시, 하이픈, 콤마, 공백 제거
          .replace(/:/g, "") // 콜론 제거
          .replace(/(\d{8})(\d{6})/, "$1_$2"); // 날짜와 시간 사이에 언더스코어 추가
        backupPath = path.join(backupDir, `${fileNameWithoutExt}_${timestamp}${fileExt}`);

        // 파일 복사
        fs.copyFileSync(filePath, backupPath);
      }

      // 파일이 열려있는지 확인
      const fileIsOpen = isFileOpen(filePath);
      
      // 모드에 따른 처리 - cells 모드 또는 full/append 모드
      let processCellData = [];
      
      if (mode === "cells") {
        // cells 모드: 사용자 제공 셀 데이터 그대로 사용
        processCellData = cellData;
      } else {
        // full 또는 append 모드: 2D 배열을 셀 데이터로 변환
        
        // append 모드면 기존 데이터 가져와서 병합
        let mergedData = [];
        
        if (mode === "append") {
          // 기존 데이터 가져오기
          let existingData = [];
          try {
            // ExcelJS를 사용하여 파일 읽기 시도
            const tempWorkbook = new ExcelJS.Workbook();
            await tempWorkbook.xlsx.readFile(filePath);

            // 시트 선택
            let tempWorksheet;
            if (sheetName) {
              tempWorksheet = tempWorkbook.getWorksheet(sheetName);
            } else if (tempWorkbook.worksheets.length > 0) {
              tempWorksheet = tempWorkbook.worksheets[0];
            }

            // 시트에서 데이터 추출
            if (tempWorksheet) {
              existingData = [];
              tempWorksheet.eachRow((row, rowNumber) => {
                const rowData = [];
                row.eachCell((cell, colNumber) => {
                  rowData[colNumber - 1] = cell.value;
                });
                existingData[rowNumber - 1] = rowData;
              });

              // 빈 행 제거 (맨 앞에 있을 수 있는 undefined 요소)
              existingData = existingData.filter((row) => row !== undefined);
            }
          } catch (err) {
            // 읽기 실패 시 빈 배열로 진행
            existingData = [];
          }

          // 데이터 병합
          if (existingData.length > 0) {
            if (smartMerge) {
              // 기존 데이터 분석
              const existingHeadersCount = existingData[0]?.length || 0;
              const newHeadersCount = data[0]?.length || 0;
              
              // 데이터 병합 방식 결정
              if (newHeadersCount > existingHeadersCount) {
                // 새 데이터에 열이 더 많은 경우: 열 확장 처리
                mergedData = existingData.map((row, rowIndex) => {
                  if (rowIndex < data.length) {
                    // 기존 데이터의 키 열을 기준으로 일치하는 행 찾기
                    // 여기서는 첫 번째 열을 키로 가정 (예: 영어단어)
                    if (row[0] === data[rowIndex][0]) {
                      // 첫 번째 열 값이 일치하면 나머지 열 확장
                      return [...row, ...data[rowIndex].slice(existingHeadersCount)];
                    }
                  }
                  return row;
                });
                
                // 기존 데이터에 없는 새 행 추가
                const existingKeys = existingData.map(row => row[0]);
                const newRows = data.filter(row => !existingKeys.includes(row[0]));
                mergedData = [...mergedData, ...newRows];
              } 
              else if (data.length > existingData.length) {
                // 새 데이터에 행이 더 많은 경우: 행 추가 처리
                const existingKeys = existingData.map(row => row[0]);
                
                // 기존 데이터와 중복되지 않는 새 행만 추가
                const newRows = data.filter(row => !existingKeys.includes(row[0]));
                mergedData = [...existingData, ...newRows];
              }
              else {
                // 동일한 구조이거나 단순 업데이트: 키 비교 후 병합
                const mergeMap = new Map();
                
                // 기존 데이터를 Map에 저장 (키: 첫 열 값)
                existingData.forEach(row => {
                  mergeMap.set(row[0], row);
                });
                
                // 새 데이터로 업데이트 또는 추가
                data.forEach(row => {
                  mergeMap.set(row[0], row);
                });
                
                // Map을 다시 배열로 변환
                mergedData = Array.from(mergeMap.values());
              }
            } else {
              // 스마트 병합 비활성화: 단순 추가
              mergedData = [...existingData, ...data];
            }
          } else {
            // 기존 데이터가 없으면 새 데이터 그대로 사용
            mergedData = data;
          }
        } else {
          // full 모드: 데이터 그대로 사용
          mergedData = data;
          
          // 현재 시트의 사용된 범위 정보 가져오기
          let maxRow = 0;
          let maxCol = 0;
          
          // 파일이 열려있지 않은 경우 기존 시트 범위 정보 가져오기
          if (!fileIsOpen) {
            try {
              const tempWorkbook = new ExcelJS.Workbook();
              await tempWorkbook.xlsx.readFile(filePath);
              const tempWorksheet = sheetName 
                ? tempWorkbook.getWorksheet(sheetName) 
                : tempWorkbook.worksheets[0];
              
              if (tempWorksheet && tempWorksheet.usedRange) {
                maxRow = tempWorksheet.rowCount || 100; // 기본값 제공
                maxCol = tempWorksheet.columnCount || 20; // 기본값 제공
              } else {
                // 기본 범위 제공
                maxRow = 100;
                maxCol = 20;
              }
            } catch (err) {
              // 읽기 실패 시 기본 범위 사용
              maxRow = 100;
              maxCol = 20;
            }
          } else {
            // 파일이 열려있는 경우 더 넓은 기본 범위 사용
            maxRow = 100;
            maxCol = 20;
          }
          
          // 명시적으로 기존 셀 데이터를 빈 값으로 설정하는 코드
          const cellDataToEmpty = [];
          // 기존 데이터의 모든 셀에 대해 빈 값 할당
          for (let row = 1; row <= maxRow; row++) {
            for (let col = 1; col <= maxCol; col++) {
              let colLetter = '';
              let tempCol = col;
              while (tempCol > 0) {
                const remainder = (tempCol - 1) % 26;
                colLetter = String.fromCharCode(65 + remainder) + colLetter;
                tempCol = Math.floor((tempCol - 1) / 26);
              }
              cellDataToEmpty.push({
                cellReference: `${colLetter}${row}`,
                value: ""
              });
            }
          }
          
          // 빈 값으로 셀 초기화 먼저 수행
          if (data.length === 0 || (data.length === 1 && data[0].length === 0)) {
            // 빈 데이터인 경우 시트 초기화만 수행
            processCellData = cellDataToEmpty;
          } else {
            // 빈 셀로 초기화 후 새 데이터 설정
            processCellData = [...cellDataToEmpty];
          }
        }
        
        // 2D 배열 데이터를 cellData 형식으로 변환 (빈 데이터가 아닌 경우만)
        if (!(mode === "full" && (data.length === 0 || (data.length === 1 && data[0].length === 0)))) {
          for (let rowIndex = 0; rowIndex < mergedData.length; rowIndex++) {
            const row = mergedData[rowIndex];
            for (let colIndex = 0; colIndex < row.length; colIndex++) {
              if (row[colIndex] === undefined || row[colIndex] === null) continue;
              
              // 열 문자 생성 (0 -> A, 1 -> B, ...)
              let colLetter = '';
              let tempCol = colIndex + 1;
              
              while (tempCol > 0) {
                const remainder = (tempCol - 1) % 26;
                colLetter = String.fromCharCode(65 + remainder) + colLetter;
                tempCol = Math.floor((tempCol - 1) / 26);
              }
              
              // A1 형식의 셀 참조 생성
              const cellRef = `${colLetter}${rowIndex + 1}`;
              
              // full 모드에서 기존에 빈 셀 초기화를 했다면 해당 셀 참조는 덮어씌우기
              if (mode === "full") {
                // 이미 초기화된 셀 참조 찾기
                const existingIndex = processCellData.findIndex(
                  cell => cell.cellReference === cellRef
                );
                
                if (existingIndex !== -1) {
                  // 기존 셀 참조 업데이트
                  processCellData[existingIndex].value = row[colIndex];
                } else {
                  // 없으면 새로 추가
                  processCellData.push({
                    cellReference: cellRef,
                    value: row[colIndex]
                  });
                }
              } else {
                // append 모드에서는 그냥 추가
                processCellData.push({
                  cellReference: cellRef,
                  value: row[colIndex]
                });
              }
            }
          }
        }
      }
      
      // 이제 모든 모드에서 셀 단위 업데이트 방식 사용
      
      if (fileIsOpen) {
        // 파일이 열려있는 경우 COM 인터페이스 사용
        const result = await updateSpecificCells(filePath, sheetName, processCellData, 3);
        
        if (result && result.success) {
          // 요청 시 Excel로 파일 열기
          if (openFile) {
            // COM 인터페이스를 통해 열린 파일을 전면에 가져옴
            const bringToFrontScript = `
            $excel = [Runtime.InteropServices.Marshal]::GetActiveObject("Excel.Application")
            $excel.Visible = $true
            $excel.Application.WindowState = -4143 # xlMaximized
            $excel.Application.Activate()
            `;
            await runPowerShellScript(bringToFrontScript);
          }

          let backupMessage = "";
          if (createBackup && backupPath) {
            backupMessage = `\n백업 생성됨: ${backupPath}`;
          }

          // 성공 개수 추출
          const successMatch = result.message.match(/총 (\d+) 개 셀이/);
          const successCount = successMatch ? successMatch[1] : processCellData.length;
          
          // 오류 개수 추출
          const errorMatch = result.message.match(/오류: (\d+) 개/);
          const errorCount = errorMatch ? parseInt(errorMatch[1]) : 0;

          // 모드에 따라 다른 성공 메시지 구성
          let successText = "";
          if (mode === "cells") {
            successText = `엑셀 파일의 특정 셀들이 성공적으로 업데이트되었습니다:
 - 파일: ${filePath}
 - 시트: ${sheetName || "활성 시트"}
 - 업데이트된 셀: ${successCount}개${errorCount > 0 ? ` (오류: ${errorCount}개)` : ''}${backupMessage}`;
          } else {
            successText = `엑셀 파일에 데이터가 성공적으로 업데이트되었습니다:
 - 파일: ${filePath}
 - 시트: ${sheetName || "활성 시트"}
 - 모드: ${mode === "append" ? (smartMerge ? "스마트 병합" : "추가") : "덮어쓰기"}
 - 업데이트된 셀: ${successCount}개${errorCount > 0 ? ` (오류: ${errorCount}개)` : ''}${backupMessage}`;
          }

          return {
            content: [
              {
                type: "text",
                text: successText + "\n\n변경 사항이 Excel에 즉시 반영되었습니다.",
              },
            ],
          };
        } else {
          // 실패했지만 파일 열기만 시도
          if (openFile) {
            await openExcelFile(filePath);
            return {
              content: [
                {
                  type: "text",
                  text: `데이터 업데이트는 실패했지만 Excel 파일을 열기 시도했습니다: ${filePath}`,
                },
              ],
            };
          }

          return {
            content: [
              {
                type: "text",
                text: `열려있는 엑셀 파일 업데이트 실패: ${
                  result ? result.message : "알 수 없는 오류"
                }`,
              },
            ],
            isError: true,
          };
        }
      } else {
        // 파일이 열려있지 않은 경우 ExcelJS 사용
        const workbook = new ExcelJS.Workbook();

        try {
          // 파일이 이미 있으면 읽기
          if (fs.existsSync(filePath)) {
            await workbook.xlsx.readFile(filePath);
          }
        } catch (err) {
          // 읽기 실패시 새 워크북으로 간주
        }

        // 시트 확인 및 생성
        let worksheet;
        if (sheetName) {
          worksheet = workbook.getWorksheet(sheetName);
          if (!worksheet) {
            worksheet = workbook.addWorksheet(sheetName);
          } else if (mode === "full") {
            // full 모드에서는 기존 데이터 지우기
            const rowCount = worksheet.rowCount;
            for (let i = rowCount; i >= 1; i--) {
              worksheet.spliceRows(i, 1);
            }
          }
        } else {
          if (workbook.worksheets.length === 0) {
            worksheet = workbook.addWorksheet("Sheet1");
          } else {
            worksheet = workbook.worksheets[0];
            if (mode === "full") {
              // full 모드에서는 기존 데이터 지우기
              const rowCount = worksheet.rowCount;
              for (let i = rowCount; i >= 1; i--) {
                worksheet.spliceRows(i, 1);
              }
            }
          }
        }

        // 셀 단위로 업데이트
        let successCount = 0;
        let errorCells = [];

        for (const cell of processCellData) {
          try {
            worksheet.getCell(cell.cellReference).value = cell.value;
            successCount++;
          } catch (err) {
            errorCells.push(cell.cellReference);
          }
        }

        // 파일 저장
        await workbook.xlsx.writeFile(filePath);

        // 요청 시 Excel로 파일 열기
        if (openFile) {
          await openExcelFile(filePath);
        }

        let backupMessage = "";
        if (createBackup && backupPath) {
          backupMessage = `\n백업 생성됨: ${backupPath}`;
        }

        // 모드에 따라 다른 성공 메시지 구성
        let successText = "";
        if (mode === "cells") {
          successText = `엑셀 파일의 특정 셀들이 성공적으로 업데이트되었습니다:
 - 파일: ${filePath}
 - 시트: ${worksheet.name}
 - 업데이트된 셀: ${successCount}개${errorCells.length > 0 ? ` (실패: ${errorCells.length}개)` : ''}${backupMessage}`;
        } else {
          successText = `엑셀 파일에 데이터가 성공적으로 업데이트되었습니다:
 - 파일: ${filePath}
 - 시트: ${worksheet.name}
 - 모드: ${mode === "append" ? (smartMerge ? "스마트 병합" : "추가") : "덮어쓰기"}
 - 업데이트된 셀: ${successCount}개${errorCells.length > 0 ? ` (실패: ${errorCells.length}개)` : ''}${backupMessage}`;
        }

        return {
          content: [
            {
              type: "text",
              text: successText,
            },
          ],
        };
      }
    } catch (error) {
      // 오류가 발생했지만 파일 열기만 시도
      if (openFile) {
        await openExcelFile(filePath);
        return {
          content: [
            {
              type: "text",
              text: `데이터 업데이트 중 오류(${error.message})가 발생했지만 Excel 파일을 열기 시도했습니다: ${filePath}`,
            },
          ],
        };
      }

      return {
        content: [
          { type: "text", text: `엑셀 데이터 업데이트 오류: ${error.message}` },
        ],
        isError: true,
      };
    }
  }
)

// 엑셀 파일 읽기 도구
server.tool(
  "read_excel",
  "엑셀 파일을 읽어 내용을 반환합니다.",
  {
    filePath: z.string().describe("읽을 엑셀 파일의 경로"),
    sheetName: z.string().optional().describe("읽을 시트 이름 (옵션)"),
    openFile: z
      .boolean()
      .optional()
      .default(false)
      .describe("파일을 Excel로 열지 여부 (기본값: false)"),
  },
  async ({ filePath, sheetName, openFile }) => {
    try {
      // 파일 존재 확인
      if (!fs.existsSync(filePath)) {
        return {
          content: [
            { type: "text", text: `파일을 찾을 수 없습니다: ${filePath}` },
          ],
          isError: true,
        };
      }

      // 파일이 열려있는지 확인
      let isFileOpen = false;
      try {
        const fd = fs.openSync(filePath, "r+");
        fs.closeSync(fd);
      } catch (e) {
        isFileOpen = true;
      }

      // 파일이 열려있다면 PowerShell을 통해 COM 인터페이스로 읽기
      if (isFileOpen) {
        try {
          // PowerShell 스크립트 작성
          const psScript = `
          try {
            # Excel 애플리케이션 객체 생성 또는 가져오기
            try {
              $excel = [Runtime.InteropServices.Marshal]::GetActiveObject("Excel.Application")
            } catch {
              $excel = New-Object -ComObject Excel.Application
            }
            
            $excel.Visible = $true
            
            # 파일 열기 시도
            $normalizedPath = "${filePath.replace(/\\/g, "\\\\")}"
            
            # 이미 열려있는 워크북 찾기
            $workbook = $null
            foreach ($wb in $excel.Workbooks) {
              if ($wb.FullName -eq $normalizedPath) {
                $workbook = $wb
                break
              }
            }
            
            # 워크북이 없으면 열기
            if ($workbook -eq $null) {
              $workbook = $excel.Workbooks.Open($normalizedPath)
            }
            
            # 시트 선택
            $worksheet = $null
            ${
              sheetName
                ? `
                try {
                  $worksheet = $workbook.Worksheets("${sheetName}")
                } catch {
                  $worksheet = $workbook.ActiveSheet
                }`
                : `
                $worksheet = $workbook.ActiveSheet`
            }
            
            # 사용된 범위 얻기
            $usedRange = $worksheet.UsedRange
            $rowCount = $usedRange.Rows.Count
            $colCount = $usedRange.Columns.Count
            
            # 시트 이름, 행 수, 열 수 기록
            $sheetInfo = "SHEET_INFO:" + $worksheet.Name + "," + $rowCount + "," + $colCount
            Write-Output $sheetInfo
            
            # 전체 워크시트의 이름 목록
            $sheetList = "SHEET_LIST:"
            foreach ($sheet in $workbook.Sheets) {
              $sheetList += $sheet.Name + ","
            }
            Write-Output $sheetList
            
            # 데이터가 없으면 여기서 종료
            if ($rowCount -eq 0 -or $colCount -eq 0) {
              Write-Output "EMPTY_SHEET"
              return
            }
            
            # 셀 데이터 읽기
            Write-Output "DATA_START"
            for ($row = 1; $row -le $rowCount; $row++) {
              $rowData = ""
              for ($col = 1; $col -le $colCount; $col++) {
                $cellValue = $usedRange.Cells.Item($row, $col).Value2
                if ($cellValue -eq $null) {
                  $cellValue = ""
                }
                $rowData += $cellValue.ToString() + "|CELL_DELIM|"
              }
              Write-Output $rowData
            }
            Write-Output "DATA_END"
            
          } catch {
            Write-Output "ERROR: $($_.Exception.Message)"
          } finally {
            # COM 객체 참조 해제 (Excel 프로그램은 종료하지 않음)
            if ($worksheet -ne $null) {
              [System.Runtime.Interopservices.Marshal]::ReleaseComObject($worksheet) | Out-Null
            }
            if ($workbook -ne $null) {
              [System.Runtime.Interopservices.Marshal]::ReleaseComObject($workbook) | Out-Null
            }
            if ($excel -ne $null) {
              [System.Runtime.Interopservices.Marshal]::ReleaseComObject($excel) | Out-Null
            }
            [System.GC]::Collect()
            [System.GC]::WaitForPendingFinalizers()
          }
          `;

          // PowerShell 스크립트 실행
          const result = await runPowerShellScript(psScript);

          // 결과 파싱
          if (result && !result.includes("ERROR:")) {
            const lines = result.split("\n");
            let data = [];
            let activeSheet = "";
            let allSheets = [];
            let rowCount = 0;
            let colCount = 0;
            let isDataSection = false;

            for (const line of lines) {
              const trimmedLine = line.trim();

              if (trimmedLine.startsWith("SHEET_INFO:")) {
                const info = trimmedLine.substring(11).split(",");
                activeSheet = info[0];
                rowCount = parseInt(info[1]) || 0;
                colCount = parseInt(info[2]) || 0;
              } else if (trimmedLine.startsWith("SHEET_LIST:")) {
                allSheets = trimmedLine
                  .substring(11)
                  .split(",")
                  .filter((s) => s.trim());
              } else if (trimmedLine === "DATA_START") {
                isDataSection = true;
                continue;
              } else if (trimmedLine === "DATA_END") {
                isDataSection = false;
              } else if (isDataSection) {
                const cells = trimmedLine.split("|CELL_DELIM|");
                cells.pop(); // 마지막 구분자 이후 빈 요소 제거
                data.push(cells);
              } else if (trimmedLine === "EMPTY_SHEET") {
                data = [];
              }
            }

            // 요청 시 Excel로 파일 열기 - 지금은 이미 열려있음
            if (openFile) {
              // 이미 열려있으므로 창만 활성화
              const activateScript = `
              try {
                $excel = [Runtime.InteropServices.Marshal]::GetActiveObject("Excel.Application")
                $excel.Visible = $true
                $excel.Application.WindowState = -4143 # xlMaximized
                $excel.Application.Activate()
              } catch {}
              `;

              await runPowerShellScript(activateScript);
            }

            // 결과 반환
            return {
              content: [
                {
                  type: "text",
                  text: `엑셀 파일 읽기 결과:
- 파일: ${filePath}
- 활성 시트: ${activeSheet}
- 전체 시트 목록: ${allSheets.join(", ")}
- 행 수: ${rowCount}
- 열 수: ${colCount}
- 데이터: ${JSON.stringify(data, null, 2)}`,
                },
              ],
            };
          } else {
            // COM 인터페이스 실패 시 기본 워크북 생성
            return {
              content: [
                {
                  type: "text",
                  text: `COM 인터페이스로 엑셀 파일을 읽을 수 없습니다: ${
                    result ? result.trim() : "알 수 없는 오류"
                  }`,
                },
              ],
              isError: true,
            };
          }
        } catch (error) {
          return {
            content: [
              {
                type: "text",
                text: `COM 인터페이스 오류: ${error.message}`,
              },
            ],
            isError: true,
          };
        }
      } else {
        // 파일이 열려있지 않은 경우 ExcelJS 사용
        const workbook = new ExcelJS.Workbook();

        try {
          // 엑셀 파일 읽기
          await workbook.xlsx.readFile(filePath);
        } catch (readErr) {
          return {
            content: [
              {
                type: "text",
                text: `엑셀 파일을 읽을 수 없습니다: ${readErr.message}`,
              },
            ],
            isError: true,
          };
        }

        // 시트 목록 가져오기
        const sheets = workbook.worksheets.map((sheet) => sheet.name);

        // 시트 선택 (지정된 시트가 없거나 존재하지 않으면 첫 번째 시트 사용)
        const selectedSheet =
          sheetName && sheets.includes(sheetName)
            ? workbook.getWorksheet(sheetName)
            : workbook.worksheets[0];

        // JSON 데이터로 변환
        const jsonData = [];
        if (selectedSheet) {
          selectedSheet.eachRow({ includeEmpty: true }, (row, rowNumber) => {
            const rowData = [];
            row.eachCell({ includeEmpty: true }, (cell, colNumber) => {
              if (colNumber > rowData.length) {
                // 중간 빈 셀 채우기
                while (rowData.length < colNumber - 1) {
                  rowData.push("");
                }
              }
              rowData.push(cell.value !== null ? cell.value : "");
            });
            jsonData.push(rowData);
          });
        }

        // 요청 시 Excel로 파일 열기
        if (openFile) {
          await openExcelFile(filePath);
        }

        return {
          content: [
            {
              type: "text",
              text: `엑셀 파일 읽기 결과:
- 파일: ${filePath}
- 활성 시트: ${selectedSheet ? selectedSheet.name : "없음"}
- 전체 시트 목록: ${sheets.join(", ")}
- 행 수: ${jsonData.length}
- 열 수: ${
                jsonData.length > 0
                  ? Math.max(...jsonData.map((row) => row.length))
                  : 0
              }
- 데이터: ${JSON.stringify(jsonData, null, 2)}`,
            },
          ],
        };
      }
    } catch (error) {
      return {
        content: [
          { type: "text", text: `엑셀 파일 읽기 오류: ${error.message}` },
        ],
        isError: true,
      };
    }
  }
);

// 시트 추가 도구
server.tool(
  "add_sheet",
  "엑셀 파일에 새 시트를 추가합니다.",
  {
    filePath: z.string().describe("엑셀 파일의 경로"),
    sheetName: z.string().describe("추가할 시트 이름"),
    data: z
      .array(z.array(z.any()))
      .optional()
      .describe("시트에 추가할 2차원 배열 형태의 데이터 (선택사항)"),
    openFile: z
      .boolean()
      .optional()
      .default(false)
      .describe("저장 후 Excel로 열지 여부 (기본값: false)"),
  },
  async ({ filePath, sheetName, data = [], openFile }) => {
    try {
      // 파일 존재 확인
      if (!fs.existsSync(filePath)) {
        return {
          content: [
            { type: "text", text: `파일을 찾을 수 없습니다: ${filePath}` },
          ],
          isError: true,
        };
      }

      // 1. 파일이 열려있는지 확인
      const fileIsOpen = isFileOpen(filePath);

      // 2. 파일이 열려있다면, PowerShell로 시트 추가
      if (fileIsOpen) {
        try {
          // PowerShell 스크립트 작성
          const psScript = `
          try {
              # Excel 애플리케이션 객체 생성 또는 가져오기
              try {
                  $excel = [Runtime.InteropServices.Marshal]::GetActiveObject("Excel.Application")
              } catch {
                  $excel = New-Object -ComObject Excel.Application
              }
              
              $excel.Visible = $true
              $excel.DisplayAlerts = $false
              
              # 파일 열기 시도
              $normalizedPath = "${filePath.replace(/\\/g, "\\\\")}"
              
              # 이미 열려있는 워크북 찾기
              $workbook = $null
              foreach ($wb in $excel.Workbooks) {
                  if ($wb.FullName -eq $normalizedPath) {
                      $workbook = $wb
                      break
                  }
              }
              
              # 워크북이 없으면 열기
              if ($workbook -eq $null) {
                  $workbook = $excel.Workbooks.Open($normalizedPath)
              }
              
              # 워크북 활성화
              $workbook.Activate()
              
              # 시트 이름 중복 검사
              $sheetExists = $false
              foreach ($sheet in $workbook.Sheets) {
                  if ($sheet.Name -eq "${sheetName}") {
                      $sheetExists = $true
                      break
                  }
              }
              
              if ($sheetExists) {
                  Write-Output "ERROR: 시트 이름 '${sheetName}'이(가) 이미 존재합니다."
                  return
              }
              
              # 새 시트 추가
              $newSheet = $workbook.Worksheets.Add()
              $newSheet.Name = "${sheetName}"
              
              # 시트 활성화
              $newSheet.Activate()
              
              # 데이터 추가 (있는 경우)
              ${
                data.length > 0
                  ? `
              # 데이터 작성
              $rowCount = ${data.length}
              $colCount = ${Math.max(
                  ...data.map((row) => row.length)
                )}
              
          # 데이터를 한 번에 설정할 범위 생성
              $targetRange = $newSheet.Range($newSheet.Cells(1, 1), $newSheet.Cells($rowCount, $colCount))
              
              # 2차원 배열 생성
              $dataArray = New-Object 'object[,]' $rowCount, $colCount
              
              ${data
                .map((row, rowIndex) => {
                  return row
                    .map((cell, colIndex) => {
                      // 값 타입에 따른 처리
                      if (typeof cell === "string") {
                        // 문자열은 따옴표로 묶고 특수문자 처리
                        const escapedValue = cell
                          .replace(/'/g, "''")
                          .replace(/"/g, '""');
                        return `$dataArray[${rowIndex}, ${colIndex}] = '${escapedValue}'`;
                      } else if (
                        cell === null ||
                        cell === undefined
                      ) {
                        // null/undefined는 빈 문자열로
                        return `$dataArray[${rowIndex}, ${colIndex}] = ''`;
                      } else if (typeof cell === "number") {
                        // 숫자는 그대로
                        return `$dataArray[${rowIndex}, ${colIndex}] = ${cell}`;
                      } else if (typeof cell === "boolean") {
                        // 불리언 값 처리
                        return `$dataArray[${rowIndex}, ${colIndex}] = $${cell}`;
                      } else {
                        // 기타 값은 문자열로 변환
                        return `$dataArray[${rowIndex}, ${colIndex}] = '${String(
                          cell
                        ).replace(/'/g, "''")}'`;
                      }
                    })
                    .join("\n");
                })
                .join("\n")}
              
              # 데이터 배열을 범위에 한 번에 설정
              $targetRange.Value2 = $dataArray
              
              # 자동 맞춤 적용 (열 너비 자동 조절)
              $newSheet.UsedRange.Columns.AutoFit() | Out-Null
              `
                  : ""
              }
              
              # 저장
              $workbook.Save()
              
              Write-Output "SUCCESS: 새 시트 '${sheetName}'이(가) 성공적으로 추가되었습니다."
          } catch {
              Write-Output "ERROR: $($_.Exception.Message)"
          } finally {
              if ($excel -ne $null) {
                  $excel.DisplayAlerts = $true
              }
          }
          `;

          // PowerShell 스크립트 실행
          const result = await runPowerShellScript(psScript);

          // 결과 확인
          if (result && result.includes("SUCCESS")) {
            // 요청 시 Excel로 파일 열기
            if (openFile) {
              await openExcelFile(filePath);
            }

            return {
              content: [{ type: "text", text: result.trim() }],
            };
          } else {
            return {
              content: [
                {
                  type: "text",
                  text: result
                    ? result.trim()
                    : "시트 추가 중 오류가 발생했습니다.",
                },
              ],
              isError: true,
            };
          }
        } catch (error) {
          return {
            content: [
              { type: "text", text: `시트 추가 오류: ${error.message}` },
            ],
            isError: true,
          };
        }
      }

      // 3. 파일이 열려있지 않은 경우 ExcelJS 사용
      // ExcelJS 워크북 생성
      const workbook = new ExcelJS.Workbook();

      try {
        // 엑셀 파일 읽기
        await workbook.xlsx.readFile(filePath);
      } catch (readErr) {
        return {
          content: [
            {
              type: "text",
              text: `엑셀 파일을 읽을 수 없습니다: ${readErr.message}`,
            },
          ],
          isError: true,
        };
      }

      // 시트 이름 중복 확인
      if (workbook.getWorksheet(sheetName)) {
        return {
          content: [
            {
              type: "text",
              text: `시트 이름 '${sheetName}'이(가) 이미 존재합니다.`,
            },
          ],
          isError: true,
        };
      }

      // 새 워크시트 추가
      const worksheet = workbook.addWorksheet(sheetName);

      // 데이터가 있는 경우 추가
      if (data.length > 0) {
        worksheet.addRows(data);
      }

      // 파일 저장
      await workbook.xlsx.writeFile(filePath);

      // 요청 시 Excel로 파일 열기
      if (openFile) {
        await openExcelFile(filePath);
      }

      return {
        content: [
          {
            type: "text",
            text: `새 시트 '${sheetName}'이(가) 성공적으로 추가되었습니다.${
              data.length > 0
                ? ` ${data.length}행의 데이터가 추가되었습니다.`
                : ""
            }`,
          },
        ],
      };
    } catch (error) {
      return {
        content: [{ type: "text", text: `시트 추가 오류: ${error.message}` }],
        isError: true,
      };
    }
  }
);

// 시트 이름 변경 도구
server.tool(
  "rename_sheet",
  "엑셀 파일에서 시트 이름을 변경합니다.",
  {
    filePath: z.string().describe("엑셀 파일의 경로"),
    currentSheetName: z.string().describe("현재 시트 이름"),
    newSheetName: z.string().describe("변경할 새 시트 이름"),
    openFile: z
      .boolean()
      .optional()
      .default(false)
      .describe("저장 후 Excel로 열지 여부 (기본값: false)"),
  },
  async ({ filePath, currentSheetName, newSheetName, openFile }) => {
    try {
      // 파일 존재 확인
      if (!fs.existsSync(filePath)) {
        return {
          content: [
            { type: "text", text: `파일을 찾을 수 없습니다: ${filePath}` },
          ],
          isError: true,
        };
      }

      // 새 시트 이름 유효성 검사
      if (newSheetName.length > 31) {
        return {
          content: [
            { type: "text", text: "시트 이름은 31자를 초과할 수 없습니다." },
          ],
          isError: true,
        };
      }

      // 시트 이름에 유효하지 않은 문자가 있는지 확인
      const invalidChars = ["/", "\\", "?", "*", "[", "]", ":", "'"];
      for (const char of invalidChars) {
        if (newSheetName.includes(char)) {
          return {
            content: [
              {
                type: "text",
                text: `시트 이름에는 다음 문자를 포함할 수 없습니다: ${invalidChars.join(
                  " "
                )}`,
              },
            ],
            isError: true,
          };
        }
      }

      // 파일이 열려있는지 확인
      const fileIsOpen = isFileOpen(filePath);

      // 파일이 열려있다면, PowerShell로 시트 이름 변경
      if (fileIsOpen) {
        try {
          // PowerShell 스크립트 작성
          const psScript = `
          try {
            # Excel 애플리케이션 객체 생성 또는 가져오기
            try {
              $excel = [Runtime.InteropServices.Marshal]::GetActiveObject("Excel.Application")
            } catch {
              $excel = New-Object -ComObject Excel.Application
            }
            
            $excel.Visible = $true
            $excel.DisplayAlerts = $false
            
            # 파일 열기 시도
            $normalizedPath = "${filePath.replace(/\\/g, "\\\\")}"
            
            # 이미 열려있는 워크북 찾기
            $workbook = $null
            foreach ($wb in $excel.Workbooks) {
              if ($wb.FullName -eq $normalizedPath) {
                $workbook = $wb
                break
              }
            }
            
            # 워크북이 없으면 열기
            if ($workbook -eq $null) {
              $workbook = $excel.Workbooks.Open($normalizedPath)
            }
            
            # 워크북 활성화
            $workbook.Activate()
            
            # 시트 찾기
            $sheetExists = $false
            $worksheet = $null
            
            foreach ($sheet in $workbook.Sheets) {
              if ($sheet.Name -eq "${currentSheetName}") {
                $sheetExists = $true
                $worksheet = $sheet
                break
              }
            }
            
            if (-not $sheetExists) {
              Write-Output "ERROR: 시트 이름 '${currentSheetName}'을(를) 찾을 수 없습니다."
              return
            }
            
            # 동일한 이름의 시트가 이미 있는지 확인
            foreach ($sheet in $workbook.Sheets) {
              if ($sheet.Name -eq "${newSheetName}") {
                Write-Output "ERROR: 시트 이름 '${newSheetName}'이(가) 이미 존재합니다."
                return
              }
            }
            
            # 시트 이름 변경
            $worksheet.Name = "${newSheetName}"
            
            # 저장
            $workbook.Save()
            
            Write-Output "SUCCESS: 시트 이름이 '${currentSheetName}'에서 '${newSheetName}'(으)로 성공적으로 변경되었습니다."
          } catch {
            Write-Output "ERROR: $($_.Exception.Message)"
          } finally {
            if ($excel -ne $null) {
              $excel.DisplayAlerts = $true
            }
          }
          `;

          // PowerShell 스크립트 실행
          const result = await runPowerShellScript(psScript);

          // 결과 확인
          if (result && result.includes("SUCCESS")) {
            // 요청 시 Excel로 파일 열기
            if (openFile) {
              await openExcelFile(filePath);
            }

            return {
              content: [{ type: "text", text: result.trim() }],
            };
          } else {
            return {
              content: [
                {
                  type: "text",
                  text: result
                    ? result.trim()
                    : "시트 이름 변경 중 오류가 발생했습니다.",
                },
              ],
              isError: true,
            };
          }
        } catch (error) {
          return {
            content: [
              { type: "text", text: `시트 이름 변경 오류: ${error.message}` },
            ],
            isError: true,
          };
        }
      }

      // 파일이 열려있지 않은 경우 ExcelJS 사용
      const workbook = new ExcelJS.Workbook();

      try {
        // 엑셀 파일 읽기
        await workbook.xlsx.readFile(filePath);
      } catch (readErr) {
        return {
          content: [
            {
              type: "text",
              text: `엑셀 파일을 읽을 수 없습니다: ${readErr.message}`,
            },
          ],
          isError: true,
        };
      }

      // 현재 시트 이름 존재 확인
      const worksheet = workbook.getWorksheet(currentSheetName);
      if (!worksheet) {
        return {
          content: [
            {
              type: "text",
              text: `시트 이름 '${currentSheetName}'을(를) 찾을 수 없습니다.`,
            },
          ],
          isError: true,
        };
      }

      // 새 시트 이름 중복 확인
      if (workbook.getWorksheet(newSheetName)) {
        return {
          content: [
            {
              type: "text",
              text: `시트 이름 '${newSheetName}'이(가) 이미 존재합니다.`,
            },
          ],
          isError: true,
        };
      }

      // 시트 이름 변경
      worksheet.name = newSheetName;

      // 파일 저장
      await workbook.xlsx.writeFile(filePath);

      // 요청 시 Excel로 파일 열기
      if (openFile) {
        await openExcelFile(filePath);
      }

      return {
        content: [
          {
            type: "text",
            text: `시트 이름이 '${currentSheetName}'에서 '${newSheetName}'(으)로 성공적으로 변경되었습니다.`,
          },
        ],
      };
    } catch (error) {
      return {
        content: [
          { type: "text", text: `시트 이름 변경 오류: ${error.message}` },
        ],
        isError: true,
      };
    }
  }
);

// 시트 삭제 도구
server.tool(
  "delete_sheet",
  "엑셀 파일에서 시트를 삭제합니다.",
  {
    filePath: z.string().describe("엑셀 파일의 경로"),
    sheetName: z.string().describe("삭제할 시트 이름"),
    openFile: z
      .boolean()
      .optional()
      .default(false)
      .describe("저장 후 Excel로 열지 여부 (기본값: false)"),
  },
  async ({ filePath, sheetName, openFile }) => {
    try {
      // 파일 존재 확인
      if (!fs.existsSync(filePath)) {
        return {
          content: [
            { type: "text", text: `파일을 찾을 수 없습니다: ${filePath}` },
          ],
          isError: true,
        };
      }

      // 파일이 열려있는지 확인
      const fileIsOpen = isFileOpen(filePath);

      // 파일이 열려있다면, PowerShell로 시트 삭제
      if (fileIsOpen) {
        try {
          // PowerShell 스크립트 작성
          const psScript = `
          try {
              # Excel 애플리케이션 객체 생성 또는 가져오기
              try {
                  $excel = [Runtime.InteropServices.Marshal]::GetActiveObject("Excel.Application")
              } catch {
                  $excel = New-Object -ComObject Excel.Application
              }
              
              $excel.Visible = $true
              $excel.DisplayAlerts = $false
              
              # 파일 열기 시도
              $normalizedPath = "${filePath.replace(/\\/g, "\\\\")}"
              
              # 이미 열려있는 워크북 찾기
              $workbook = $null
              foreach ($wb in $excel.Workbooks) {
                  if ($wb.FullName -eq $normalizedPath) {
                      $workbook = $wb
                      break
                  }
              }
              
              # 워크북이 없으면 열기
              if ($workbook -eq $null) {
                  $workbook = $excel.Workbooks.Open($normalizedPath)
              }
              
              # 워크북 활성화
              $workbook.Activate()
              
              # 시트가 존재하는지 확인
              $sheetExists = $false
              $worksheet = $null
              
              foreach ($sheet in $workbook.Sheets) {
                  if ($sheet.Name -eq "${sheetName}") {
                      $sheetExists = $true
                      $worksheet = $sheet
                      break
                  }
              }
              
              if (-not $sheetExists) {
                  Write-Output "ERROR: 시트 이름 '${sheetName}'을(를) 찾을 수 없습니다."
                  return
              }
              
              # 시트가 마지막 시트인지 확인 (Excel은 항상 최소 1개의 시트가 있어야 함)
              if ($workbook.Sheets.Count -eq 1) {
                  Write-Output "ERROR: 마지막 남은 시트는 삭제할 수 없습니다. Excel 파일에는 최소 하나의 시트가 있어야 합니다."
                  return
              }
              
              # 시트 삭제
              $worksheet.Delete()
              
              # 저장
              $workbook.Save()
              
              Write-Output "SUCCESS: 시트 '${sheetName}'이(가) 성공적으로 삭제되었습니다."
          } catch {
              Write-Output "ERROR: $($_.Exception.Message)"
          } finally {
              if ($excel -ne $null) {
                  $excel.DisplayAlerts = $true
              }
          }
          `;

          // PowerShell 스크립트 실행
          const result = await runPowerShellScript(psScript);

          // 결과 확인
          if (result && result.includes("SUCCESS")) {
            // 요청 시 Excel로 파일 열기
            if (openFile) {
              await openExcelFile(filePath);
            }

            return {
              content: [{ type: "text", text: result.trim() }],
            };
          } else {
            return {
              content: [
                {
                  type: "text",
                  text: result
                    ? result.trim()
                    : "시트 삭제 중 오류가 발생했습니다.",
                },
              ],
              isError: true,
            };
          }
        } catch (error) {
          return {
            content: [
              { type: "text", text: `시트 삭제 오류: ${error.message}` },
            ],
            isError: true,
          };
        }
      }

      // 파일이 열려있지 않은 경우 ExcelJS 사용
      const workbook = new ExcelJS.Workbook();

      try {
        // 엑셀 파일 읽기
        await workbook.xlsx.readFile(filePath);
      } catch (readErr) {
        return {
          content: [
            {
              type: "text",
              text: `엑셀 파일을 읽을 수 없습니다: ${readErr.message}`,
            },
          ],
          isError: true,
        };
      }

      // 시트 존재 확인
      const worksheet = workbook.getWorksheet(sheetName);
      if (!worksheet) {
        return {
          content: [
            {
              type: "text",
              text: `시트 이름 '${sheetName}'을(를) 찾을 수 없습니다.`,
            },
          ],
          isError: true,
        };
      }

      // 워크북에 시트가 하나만 있는지 확인
      if (workbook.worksheets.length === 1) {
        return {
          content: [
            {
              type: "text",
              text: "마지막 남은 시트는 삭제할 수 없습니다. Excel 파일에는 최소 하나의 시트가 있어야 합니다.",
            },
          ],
          isError: true,
        };
      }

      // 시트 삭제
      workbook.removeWorksheet(worksheet.id);

      // 파일 저장
      await workbook.xlsx.writeFile(filePath);

      // 요청 시 Excel로 파일 열기
      if (openFile) {
        await openExcelFile(filePath);
      }

      return {
        content: [
          {
            type: "text",
            text: `시트 '${sheetName}'이(가) 성공적으로 삭제되었습니다.`,
          },
        ],
      };
    } catch (error) {
      return {
        content: [{ type: "text", text: `시트 삭제 오류: ${error.message}` }],
        isError: true,
      };
    }
  }
);



// 열려있는 Excel 파일 목록 조회 도구
server.tool(
  "list_open_excel_files",
  "현재 PC에서 열려있는 모든 Excel 파일 목록을 조회합니다.",
  {
    details: z
      .boolean()
      .optional()
      .default(false)
      .describe("파일 세부 정보 포함 여부 (기본값: false)"),
  },
  async ({ details }) => {
    try {
      // PowerShell 스크립트 작성
      const psScript = `
      try {
        # Excel 프로세스 확인
        $excelProcesses = Get-Process -Name "excel" -ErrorAction SilentlyContinue
        
        if ($null -eq $excelProcesses -or $excelProcesses.Count -eq 0) {
          Write-Output "NO_EXCEL_RUNNING"
          exit
        }
        
        Write-Output "EXCEL_RUNNING:$($excelProcesses.Count)"
        
        # COM 객체 생성 시도
        try {
          $excel = [Runtime.InteropServices.Marshal]::GetActiveObject("Excel.Application")
          Write-Output "COM_SUCCESS"
          
          # 열린 통합 문서 확인
          if ($excel.Workbooks.Count -eq 0) {
            Write-Output "NO_OPEN_WORKBOOKS"
            exit
          }
          
          Write-Output "WORKBOOK_COUNT:$($excel.Workbooks.Count)"
          
          # 각 통합 문서 정보 출력
          foreach ($wb in $excel.Workbooks) {
            Write-Output "WORKBOOK_PATH:$($wb.FullName)"
            Write-Output "WORKBOOK_NAME:$($wb.Name)"
            Write-Output "SAVED_STATUS:$($wb.Saved)"
            Write-Output "ACTIVE_SHEET:$($wb.ActiveSheet.Name)"
            
            if ($${details}) {
              # 시트 목록
              Write-Output "SHEETS_START"
              foreach ($ws in $wb.Worksheets) {
                Write-Output "SHEET:$($ws.Name)"
              }
              Write-Output "SHEETS_END"
              
              # 추가 정보
              try {
                $lastAuthor = $wb.BuiltinDocumentProperties.Item("Last Author").Value
                Write-Output "LAST_AUTHOR:$lastAuthor"
              } catch {
                Write-Output "LAST_AUTHOR:Unknown"
              }
              
              try {
                $lastSaved = $wb.BuiltinDocumentProperties.Item("Last Save Time").Value
                Write-Output "LAST_SAVED:$lastSaved"
              } catch {
                Write-Output "LAST_SAVED:Unknown"
              }
            }
            
            Write-Output "WORKBOOK_END"
          }
        } catch {
          Write-Output "COM_FAILED:$($_.Exception.Message)"
          
          # COM 접근 실패 시 프로세스 정보만으로 파일 추정
          foreach ($proc in $excelProcesses) {
            $mainWindowTitle = $proc.MainWindowTitle
            if (-not [string]::IsNullOrEmpty($mainWindowTitle) -and $mainWindowTitle -ne "Microsoft Excel") {
              # 제목에서 "(Compatibility Mode)" 및 "- Excel" 제거
              $title = $mainWindowTitle -replace " \(Compatibility Mode\)", "" -replace " - Excel$", ""
              Write-Output "WINDOW_TITLE:$title"
            }
          }
        }
      } catch {
        Write-Output "ERROR:$($_.Exception.Message)"
      } finally {
        # 리소스 해제
        if ($null -ne $excel) {
          [System.Runtime.InteropServices.Marshal]::ReleaseComObject($excel) | Out-Null
        }
        [System.GC]::Collect()
        [System.GC]::WaitForPendingFinalizers()
      }
      `;

      // PowerShell 스크립트 실행
      const result = await runPowerShellScript(psScript);

      // 결과 파싱 및 정보 추출
      const lines = result
        .split("\n")
        .map((line) => line.trim())
        .filter((line) => line);
      const openFiles = [];
      let currentFile = null;
      let excelRunning = false;
      let workbookCount = 0;
      let comSuccess = false;

      for (const line of lines) {
        if (line === "NO_EXCEL_RUNNING") {
          return {
            content: [
              { type: "text", text: "현재 Excel이 실행되고 있지 않습니다." },
            ],
          };
        } else if (line.startsWith("EXCEL_RUNNING:")) {
          excelRunning = true;
          const processCount =
            parseInt(line.substring("EXCEL_RUNNING:".length), 10) || 0;
        } else if (line === "COM_SUCCESS") {
          comSuccess = true;
        } else if (line === "NO_OPEN_WORKBOOKS") {
          return {
            content: [
              {
                type: "text",
                text: "Excel이 실행 중이지만 열려있는 통합 문서가 없습니다.",
              },
            ],
          };
        } else if (line.startsWith("WORKBOOK_COUNT:")) {
          workbookCount =
            parseInt(line.substring("WORKBOOK_COUNT:".length), 10) || 0;
        } else if (line.startsWith("WORKBOOK_PATH:")) {
          // 새 통합 문서 정보 시작
          if (currentFile) {
            openFiles.push(currentFile);
          }
          currentFile = {
            path: line.substring("WORKBOOK_PATH:".length),
            name: "",
            saved: true,
            activeSheet: "",
            sheets: [],
            lastAuthor: "",
            lastSaved: "",
          };
        } else if (line.startsWith("WORKBOOK_NAME:") && currentFile) {
          currentFile.name = line.substring("WORKBOOK_NAME:".length);
        } else if (line.startsWith("SAVED_STATUS:") && currentFile) {
          currentFile.saved =
            line.substring("SAVED_STATUS:".length).toLowerCase() === "true";
        } else if (line.startsWith("ACTIVE_SHEET:") && currentFile) {
          currentFile.activeSheet = line.substring("ACTIVE_SHEET:".length);
        } else if (line.startsWith("SHEET:") && currentFile) {
          currentFile.sheets.push(line.substring("SHEET:".length));
        } else if (line.startsWith("LAST_AUTHOR:") && currentFile) {
          currentFile.lastAuthor = line.substring("LAST_AUTHOR:".length);
        } else if (line.startsWith("LAST_SAVED:") && currentFile) {
          currentFile.lastSaved = line.substring("LAST_SAVED:".length);
        } else if (line === "WORKBOOK_END" && currentFile) {
          openFiles.push(currentFile);
          currentFile = null;
        } else if (line.startsWith("WINDOW_TITLE:")) {
          // COM 접근이 실패한 경우 창 제목에서 추정
          openFiles.push({
            path: "알 수 없음",
            name: line.substring("WINDOW_TITLE:".length),
            saved: false,
            activeSheet: "알 수 없음",
            sheets: [],
            lastAuthor: "",
            lastSaved: "",
          });
        } else if (line.startsWith("COM_FAILED:")) {
          // COM 실패 - 제한된 정보 수집 시도 중
        } else if (line.startsWith("ERROR:")) {
          return {
            content: [
              {
                type: "text",
                text: `Excel 정보 조회 중 오류가 발생했습니다: ${line.substring(
                  "ERROR:".length
                )}`,
              },
            ],
            isError: true,
          };
        }
      }

      // 마지막 파일 추가
      if (currentFile) {
        openFiles.push(currentFile);
      }

      // 결과 반환
      if (openFiles.length === 0) {
        if (excelRunning) {
          return {
            content: [
              {
                type: "text",
                text: "Excel이 실행 중이지만 열린 파일 정보를 확인할 수 없습니다.",
              },
            ],
          };
        } else {
          return {
            content: [
              { type: "text", text: "열려있는 Excel 파일이 없습니다." },
            ],
          };
        }
      }

      // 결과 포맷팅
      let responseText = `현재 ${openFiles.length}개의 Excel 파일이 열려 있습니다:\n\n`;

      openFiles.forEach((file, index) => {
        responseText += `${index + 1}. ${file.name}${
          !file.saved ? " *" : ""
        }\n`;
        responseText += `   경로: ${file.path}\n`;
        responseText += `   활성 시트: ${file.activeSheet}\n`;

        if (details && file.sheets.length > 0) {
          responseText += `   시트 목록: ${file.sheets.join(", ")}\n`;
        }

        if (details && file.lastAuthor) {
          responseText += `   마지막 편집자: ${file.lastAuthor}\n`;
        }

        if (details && file.lastSaved) {
          responseText += `   마지막 저장 시간: ${file.lastSaved}\n`;
        }

        responseText += "\n";
      });

      return {
        content: [{ type: "text", text: responseText }],
      };
    } catch (error) {
      return {
        content: [
          {
            type: "text",
            text: `열린 Excel 파일 목록 조회 중 오류가 발생했습니다: ${error.message}`,
          },
        ],
        isError: true,
      };
    }
  }
);

// Excel 파일 닫기 도구
server.tool(
  "close_excel_file",
  "열려있는 Excel 파일을 닫습니다.",
  {
    filePath: z
      .string()
      .optional()
      .describe("닫을 Excel 파일의 경로 (지정하지 않으면 모든 파일)"),
    saveChanges: z
      .boolean()
      .optional()
      .default(true)
      .describe("변경사항 저장 여부 (기본값: true)"),
    closeExcel: z
      .boolean()
      .optional()
      .default(false)
      .describe("Excel 프로그램까지 종료할지 여부 (기본값: false)"),
  },
  async ({ filePath, saveChanges, closeExcel }) => {
    try {
      // PowerShell 스크립트 작성
      let psScript = `
      try {
        $excelWasRunning = $false
        
        # Excel 애플리케이션 객체 가져오기 시도
        try {
          $excel = [Runtime.InteropServices.Marshal]::GetActiveObject("Excel.Application")
          $excelWasRunning = $true
        } catch {
          # Excel이 실행 중이 아니면 종료
          Write-Output "NO_EXCEL: Excel이 실행 중이지 않습니다."
          exit
        }
        
        $excel.DisplayAlerts = $false
        
        # 파일 이름 목록 수집 (출력용)
        $fileNames = @()
        $fileCount = 0
        
        if ($excel.Workbooks.Count -gt 0) {
          $fileCount = $excel.Workbooks.Count
          foreach ($wb in $excel.Workbooks) {
            $fileNames += $wb.Name
          }
        } else {
          Write-Output "NO_WORKBOOKS: 열려있는 통합 문서가 없습니다."
        }
      `;

      if (filePath) {
        // 특정 파일 지정 시에도 파일 이름만으로 찾는 방식 추가
        psScript += `
        # 통합 문서가 있는 경우에만 파일 닫기 시도
        if ($excel.Workbooks.Count -gt 0) {
          $targetFileName = [System.IO.Path]::GetFileName("${filePath.replace(
            /\\/g,
            "\\\\"
          )}")
          $normalizedPath = "${filePath.replace(/\\/g, "\\\\")}"
          $workbookFound = $false
          
          # 이름 또는 전체 경로로 파일 찾기
          foreach ($wb in $excel.Workbooks) {
            # 파일 이름이나 전체 경로가 일치하는지 확인
            if ($wb.Name -eq $targetFileName -or $wb.FullName -eq $normalizedPath) {
              # 변경사항 저장 여부에 따라 처리
              if ($${saveChanges.toString()}) {
                $wb.Save()
              }
              $wb.Close($${saveChanges.toString()})
              $workbookFound = $true
              Write-Output "FILE_CLOSED: 파일을 성공적으로 닫았습니다: $($wb.Name)"
              break
            }
          }
          
          if (-not $workbookFound) {
            # 파일을 찾지 못했을 때 모든 파일 목록 제공
            Write-Output "FILE_NOT_FOUND: 지정한 파일을 찾을 수 없습니다."
          }
        }
        `;
      } else {
        // 모든 파일 닫기
        psScript += `
        # 통합 문서가 있는 경우에만 파일 닫기 시도
        if ($excel.Workbooks.Count -gt 0) {
          # 모든 통합 문서 닫기
          while ($excel.Workbooks.Count -gt 0) {
            $wb = $excel.Workbooks.Item(1)
            if ($${saveChanges.toString()}) {
              $wb.Save()
            }
            $wb.Close($${saveChanges.toString()})
          }
          
          Write-Output "ALL_CLOSED: $fileCount 개의 파일을 닫았습니다: $($fileNames -join ', ')"
        }
        `;
      }

      // Excel 프로그램 종료 옵션 개선 - 항상 종료 시도
      psScript += `
        # Excel 종료 옵션이 true이거나 Excel은 실행 중이지만 열린 문서가 없는 경우
        if ($${closeExcel.toString()} -or ($excelWasRunning -and $excel.Workbooks.Count -eq 0)) {
          $excel.Quit()
          Write-Output "EXCEL_CLOSED: Excel 프로그램을 종료했습니다."
          
          # 추가: Excel 프로세스 종료 시도 - 더 확실하게 종료하기 위함
          try {
            Get-Process excel -ErrorAction SilentlyContinue | Stop-Process -Force
            Write-Output "EXCEL_PROCESS_KILLED: Excel 프로세스를 강제 종료했습니다."
          } catch {
            # 프로세스 종료 실패는 무시
          }
        }
      } catch {
        Write-Output "ERROR: $($_.Exception.Message)"
      } finally {
        if ($excel -ne $null) {
          $excel.DisplayAlerts = $true
          [System.Runtime.Interopservices.Marshal]::ReleaseComObject($excel) | Out-Null
        }
        [System.GC]::Collect()
        [System.GC]::WaitForPendingFinalizers()
      }
      `;

      // PowerShell 스크립트 실행
      const result = await runPowerShellScript(psScript);

      // 결과 파싱 및 응답 생성
      const lines = result
        .split("\n")
        .map((line) => line.trim())
        .filter((line) => line);

      // Excel 닫기 결과를 저장할 변수들
      let successMessage = "";
      let errorMessage = "";
      let excelClosed = false;

      for (const line of lines) {
        if (line.startsWith("NO_EXCEL:")) {
          return {
            content: [{ type: "text", text: "Excel이 실행 중이지 않습니다." }],
          };
        } else if (line.startsWith("NO_WORKBOOKS:")) {
          // 열린 문서가 없는 경우도 기록
          successMessage =
            "Excel은 실행 중이지만 열려있는 통합 문서가 없습니다.";
        } else if (line.startsWith("FILE_CLOSED:")) {
          successMessage = line.substring("FILE_CLOSED:".length).trim();
        } else if (line.startsWith("FILE_NOT_FOUND:")) {
          errorMessage = line.substring("FILE_NOT_FOUND:".length).trim();
        } else if (line.startsWith("ALL_CLOSED:")) {
          successMessage = line.substring("ALL_CLOSED:".length).trim();
        } else if (
          line.startsWith("EXCEL_CLOSED:") ||
          line.startsWith("EXCEL_PROCESS_KILLED:")
        ) {
          excelClosed = true;
        } else if (line.startsWith("ERROR:")) {
          errorMessage = `Excel 파일 닫기 오류: ${line
            .substring("ERROR:".length)
            .trim()}`;
        }
      }

      // 최종 메시지 구성
      if (excelClosed) {
        if (successMessage) {
          successMessage += " Excel 프로그램도 완전히 종료되었습니다.";
        } else {
          successMessage = "Excel 프로그램이 완전히 종료되었습니다.";
        }
      }

      // 결과 반환
      if (successMessage) {
        return {
          content: [{ type: "text", text: successMessage }],
        };
      } else if (errorMessage) {
        return {
          content: [{ type: "text", text: errorMessage }],
          isError: true,
        };
      } else {
        // 기본 성공 메시지
        return {
          content: [{ type: "text", text: "Excel 작업이 완료되었습니다." }],
        };
      }
    } catch (error) {
      return {
        content: [
          {
            type: "text",
            text: `Excel 파일 닫기 중 오류가 발생했습니다: ${error.message}`,
          },
        ],
        isError: true,
      };
    }
  }
);

// 서버 연결
const transport = new StdioServerTransport();

// 서버와 전송 계층 연결
server.connect(transport);

// 예기치 않은 오류 처리
process.on("uncaughtException", (err) => {
  console.error("예기치 않은 오류 발생:", err);
  // 심각한 오류 발생 시 로깅만 하고 프로세스는 계속 실행
});

// 처리되지 않은 Promise 거부 처리
process.on("unhandledRejection", (reason, promise) => {
  console.error("처리되지 않은 Promise 거부:", reason);
});

console.error("Excel Controller 서버가 시작되었습니다.");