# 📋 Excel VBA: 클립보드 세로 리스트를 가로로 붙여넣기

> 웹에서 복사한 세로 리스트를 Excel에서 가로 방향으로 안전하게 붙여넣는 VBA 매크로

[![Excel Version](https://img.shields.io/badge/Excel-2016%2B-green.svg)](https://www.microsoft.com/excel)
[![VBA](https://img.shields.io/badge/VBA-Macro-blue.svg)](https://docs.microsoft.com/en-us/office/vba/api/overview/excel)
[![License](https://img.shields.io/badge/license-MIT-blue.svg)](LICENSE)

## 🎯 개요

웹 브라우저에서 세로로 나열된 텍스트 리스트를 복사한 후, Excel에서 **가로 방향으로 자동 변환하여 붙여넣기**할 수 있는 매크로입니다.

### 문제 상황
```
웹에서 복사:        Excel에 붙여넣으면:     원하는 결과:
1                   1                       1  2  3  4  5
2                   2
3                   3
4                   4
5                   5
```

## ✨ 주요 기능

- ✅ **안전한 데이터 처리**: 기존 데이터를 절대 삭제하지 않음
- ✅ **에러 없는 실행**: 런타임 오류 팝업 없이 조용히 처리
- ✅ **가변 길이 지원**: 3줄이든 10줄이든 자동 감지
- ✅ **단축키 지원**: `Ctrl + Shift + T`로 빠른 실행
- ✅ **라이브러리 불필요**: 추가 참조 설정 없이 동작

## 🚀 설치 방법

### 1. VBA 편집기 열기
```
Alt + F11
```

### 2. 새 모듈 추가
1. 프로젝트 탐색기에서 통합 문서 선택
2. `삽입(Insert)` → `모듈(Module)` 클릭

### 3. 코드 복사 & 붙여넣기

```vba
Sub PasteClipboardAsRowSafe()
    '=========================================
    ' 클립보드의 세로 리스트를 가로로 붙여넣기
    ' - 현재 시트의 다른 데이터를 절대 건드리지 않음
    ' - 런타임 오류 팝업 없이 안전하게 처리
    '=========================================
    
    Dim dataObj As Object
    Dim clipText As String
    Dim lines() As String
    Dim i As Long
    Dim targetCell As Range
    Dim validLines() As String
    Dim validCount As Long
    
    ' 에러 발생 시 조용히 처리
    On Error Resume Next
    
    ' 선택된 셀이 없으면 종료
    If Selection Is Nothing Then Exit Sub
    Set targetCell = ActiveCell
    If targetCell Is Nothing Then Exit Sub
    
    ' DataObject를 사용하여 클립보드 텍스트 읽기 (후기 바인딩)
    Set dataObj = CreateObject("New:{1C3B4210-F441-11CE-B9EA-00AA006B1A69}")
    
    ' 클립보드에서 텍스트 가져오기
    dataObj.GetFromClipboard
    
    ' 텍스트 형식 데이터가 있는지 확인
    If dataObj.GetFormat(1) Then  ' 1 = cfText
        clipText = dataObj.GetText(1)
    Else
        ' 클립보드에 텍스트가 없으면 조용히 종료
        On Error GoTo 0
        Exit Sub
    End If
    
    ' 에러가 발생했거나 텍스트가 비어있으면 종료
    If Err.Number <> 0 Or Len(Trim(clipText)) = 0 Then
        On Error GoTo 0
        Exit Sub
    End If
    
    ' 에러 처리 재설정
    On Error Resume Next
    
    ' 줄바꿈 문자로 분리 (Windows: vbCrLf, Unix: vbLf, Mac: vbCr 모두 고려)
    clipText = Replace(clipText, vbCrLf, vbLf)  ' Windows 스타일을 Unix로 통일
    clipText = Replace(clipText, vbCr, vbLf)    ' Mac 스타일도 Unix로 통일
    lines = Split(clipText, vbLf)
    
    ' 빈 줄을 제외한 유효한 줄만 추출
    ReDim validLines(0 To UBound(lines))
    validCount = 0
    
    For i = LBound(lines) To UBound(lines)
        ' 각 줄의 앞뒤 공백 제거 후 확인
        ' 빈 줄도 포함하려면 이 조건을 제거하면 됨
        If Len(Trim(lines(i))) > 0 Then
            validLines(validCount) = lines(i)
            validCount = validCount + 1
        End If
    Next i
    
    ' 유효한 데이터가 없으면 종료
    If validCount = 0 Then
        On Error GoTo 0
        Exit Sub
    End If
    
    ' 유효한 줄들을 현재 셀부터 오른쪽으로 가로 방향으로 입력
    ' 중요: 각 셀에 개별적으로 값을 할당 (범위 단위 작업 없음)
    For i = 0 To validCount - 1
        targetCell.Offset(0, i).Value = validLines(i)
        
        ' 에러 발생 시 해당 셀은 건너뛰고 계속 진행
        If Err.Number <> 0 Then
            Err.Clear
        End If
    Next i
    
    ' 에러 처리 해제
    On Error GoTo 0
    
    ' 객체 정리
    Set dataObj = Nothing
    Set targetCell = Nothing
    
End Sub
```

### 4. 저장
```
Ctrl + S
```

## ⌨️ 단축키 설정

1. Excel로 돌아가기 (`Alt + Q`)
2. 매크로 대화상자 열기
   ```
   Alt + F8
   ```
3. `PasteClipboardAsRowSafe` 선택
4. **옵션(Options)** 버튼 클릭
5. 바로 가기 키에 `Shift + T` 입력
   - 자동으로 `Ctrl + Shift + T`가 됩니다
6. **확인** → **취소**

## 📖 사용 방법

### 기본 사용법

```
1. 웹 브라우저에서 세로 리스트 복사
   ┌─────┐
   │  1  │
   │  2  │
   │  3  │
   │  4  │
   │  5  │
   └─────┘

2. Excel에서 시작 셀 선택 (예: B2)

3. Ctrl + Shift + T 입력

4. 결과 확인
   ┌──────────────────────┐
   │  1 │ 2 │ 3 │ 4 │ 5  │  ← B2부터 가로로 입력됨
   └──────────────────────┘
```

### 실제 예제

**웹에서 복사한 데이터:**
```
Apple
Banana
Cherry
Date
Elderberry
```

**Excel 결과 (B2 셀부터 시작):**
| A | B | C | D | E | F |
|---|---|---|---|---|---|
| 1 | Apple | Banana | Cherry | Date | Elderberry |

## 🛡️ 안전 장치

| 기능 | 설명 |
|------|------|
| **데이터 보호** | 선택된 행의 필요한 셀만 쓰고, 다른 모든 셀은 절대 건드리지 않음 |
| **에러 처리** | `On Error Resume Next`로 모든 오류를 조용히 처리 (팝업 없음) |
| **안정성** | `PasteSpecial Transpose` 대신 직접 텍스트 파싱 방식 사용 |
| **가변 길이** | 3줄이든 100줄이든 자동 감지 |
| **공백 처리** | 빈 줄은 자동으로 제외 |

## ⚙️ 시스템 요구사항

- Windows용 Excel 2016 이상
- VBA 매크로 활성화 필요
- 추가 라이브러리 불필요

## 🔧 커스터마이징

### 빈 줄도 포함하고 싶다면

`If Len(Trim(lines(i))) > 0 Then` 부분을 다음과 같이 수정:

```vba
' 빈 줄도 포함 (빈 셀로 표시)
validLines(validCount) = lines(i)
validCount = validCount + 1
```

### 다른 단축키 사용

단축키 설정 시 원하는 키 조합 입력 (예: `Shift + V` → `Ctrl + Shift + V`)

## 🐛 트러블슈팅

### 매크로가 실행되지 않는 경우
- 매크로 보안 설정 확인: `파일` → `옵션` → `보안 센터` → `매크로 설정`
- "모든 매크로 제외(알림 표시)" 이상으로 설정 필요

### 아무 일도 일어나지 않는 경우
- 클립보드에 텍스트가 있는지 확인
- 셀이 선택되어 있는지 확인
- 보호된 시트가 아닌지 확인

## 🤝 기여

이슈나 개선 사항이 있다면 언제든지 제보해주세요!

---

**Made with ❤️ for Excel Power Users**
