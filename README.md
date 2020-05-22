# UiPath invoke VBA 2way(Excel Macro, Invoke VBA)  :microscope:

##### UiPath에서 VBA를 활용하는 2가지 방법으로 나누어져 있다. 어떤 것을 활용할지에 대해서 아래의 2가지 방법을 보고 판단하면 된다.  
  
  
<br><br>

1. Excel Macro Activity 활용
```
- 이건 VBA가 이미 입력이 되어 있는 Excel에서 활용이 가능하다.
- 기존 엑셀에 입력된 VBA를 활용하기 때문에 엑셀 내부의 Alt + F11 를 통해서 
활용할 VBA 소스의 함수를 기억하고 UiPath Studio 내에 사용할 메크로 함수명을 입력하면 실행된다.
- Excel Macro Activity는 항상 Excel Application Scope에 포함되어야 한다.
- 엑셀에 이미 Macro형태가 저장되어 있는 형태이기 때문에 엑셀의 확장자명이 .xlsm이어야 한다.
```
<br>

2. invoke VBA Activity 활용
```
- 비교적 쓰임이 자유롭다. 
- txt 파일에 VBA Script를 입력하고 Studio 내에서 적용할 Excel 파일과 VBA 파일내에 있는 함수명을 입력한다.
- invoke VBA Activity 또한 Excel Application Scope 내에 포함되어야한다.
- 엑셀에서 외부 메크로 사용을 허용해줘야 실행된다.
(설정 -> 옵션 -> 보안센터 -> 보안센터 설정 -> 매크로 설정 -> VBA 프로젝트 객체 모델에 엑세스 체크)
- 주의점은 이미 입력된 메크로와 txt에서 불러올 메크로 중복이 되면 충돌이 난다.
```
<br><br>
##### RPA를 활용하여 다수의 Excel 파일에 메크로를 적용해보자 :rocket:

실행이 되지 않는다면 엑셀 파일을 열어보고 문서간의 충돌도 확인해보고 작업중인 엑셀 프로세스 종료하고 다시 시도해본다.



#### how to use?
https://labft3231.github.io/rpa/update/2020/03/20/UiPath-invoke-VBA-Script/
