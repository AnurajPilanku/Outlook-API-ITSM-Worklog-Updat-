*** Settings ***
Test Teardown     Run Keyword If Test Failed    update automationcenter failure    ${TEST MESSAGE}
Library           AutomationCenter
Library           Collections
Library           lib/monitormail_update_itsm.py

*** Variables ***
#input
${aeuser}         ${EMPTY}
${aeparameters}    ${EMPTY}
${aeci}           ${EMPTY}
${aepassword}     ${EMPTY}
${aeprivatekey}    ${EMPTY}
${aedatastore}    ${EMPTY}
${timeout}        NONE
#result
${output}         ""
${error}          ""
${rc}             0

*** Test Cases ***
run
    [Timeout]    ${timeout}
    &{aeparameters}=    AutomationCenter.json_parse    ${aeparameters}
    &{aeci}=    AutomationCenter.json_parse    ${aeci}
    &{datastore}=    AutomationCenter.json_parse    ${aedatastore}
    ${output}=    postconnect
    ${result}=    update automationcenter    ${output}    ${error}    ${rc}
