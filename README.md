## 1. 요구사항 정리

**핵심 기능:**
- 클립보드의 세로 리스트(줄바꿈 구분)를 현재 선택된 셀부터 오른쪽으로 가로 방향으로 붙여넣기

**절대 지켜야 할 제약:**
- ✅ **건드리면 안 되는 범위**: 선택된 셀의 행에서, 붙여넣을 셀 범위를 제외한 모든 셀(특히 같은 열의 아래쪽 행 데이터)
- ✅ **에러 처리**: 런타임 오류 팝업이 절대 표시되지 않도록 `On Error` 구문으로 안전하게 처리
- ✅ **구현 방식**: `PasteSpecial Transpose` 방식 대신, 클립보드 텍스트를 직접 읽어서 `Split`으로 처리 (B안 채택)

## 2. 완전한 VBA 코드

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

## 3. 설치 및 설정 방법

### 📋 Step 1: VBA 편집기 열기
1. Excel을 엽니다
2. `Alt + F11` 키를 눌러 VBA 편집기를 엽니다

### 📋 Step 2: 새 모듈 추가
1. VBA 편집기 왼쪽의 프로젝트 탐색기에서 해당 통합 문서를 선택
2. 메뉴 `삽입(Insert)` → `모듈(Module)` 클릭
3. 새 모듈 창에 위 코드 전체를 붙여넣기
4. `Ctrl + S` 저장

### 📋 Step 3: 매크로에 단축키 연결
1. Excel로 돌아오기 (`Alt + Q`)
2. `Alt + F8` → 매크로 목록에서 `PasteClipboardAsRowSafe` 선택
3. "옵션" 클릭 → 바로 가기 키에 `Shift + T` 입력 (결과적으로 Ctrl + Shift + T)
4. 확인 후 닫기

### 📋 Step 4: 사용 방법
1. 웹/문서 등에서 세로로 정리된 목록 복사 (예: 1 ↵ 2 ↵ 3 ↵ 4)
2. Excel에서 시작 셀 하나 선택
3. `Ctrl + Shift + T` 실행
4. 선택 셀부터 오른쪽으로 값이 가로로 채워짐

---

## ✅ 안전 장치 확인

1. 데이터 보호: 선택된 행의 필요한 셀만 씀 (다른 행/열은 영향 없음)
2. 에러 억제: 모든 런타임 오류는 조용히 무시
3. 안정성: PasteSpecial Transpose 미사용, 텍스트 직접 처리
4. 가변 길이 대응: 줄 수 자동 인식
5. 빈 줄 제거: 공백만 있는 줄 필터링

문제가 발생하거나 기능 확장이 필요하면 언제든지 요청 주세요! 🚀
