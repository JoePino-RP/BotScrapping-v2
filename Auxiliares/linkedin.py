import Functions_Bot as FB

FB.LoginELEM("1","1")
FB.D_CHR.get("https://www.linkedin.com/uas/login?session_redirect=https%3A%2F%2Fwww%2Elinkedin%2Ecom%2Fsearch%2Fresults%2Fpeople%2F%3FgeoUrn%3D%255B%2522100876405%2522%255D%26keywords%3Dtecnologia%26origin%3DGLOBAL_SEARCH_HEADER%26sid%3Dm_f&fromSignIn=true&trk=cold_join_sign_in")
usLink =FB.D_CHR.find_element_by_id("username")
usLink.send_keys("soyexbot@gmail.com")
psLink = FB.D_CHR.find_element_by_id("password")
psLink.send_keys("Marketing2022.")
psLink.send_keys(FB.Keys.ENTER)

FB.D_CHR.find_element_by_xpath("//div[@id='global-nav-typeahead']//input").clear()

busc= FB.D_CHR.find_element_by_xpath("//div[@id='global-nav-typeahead']//input")
busc.send_keys("Desarrollador Frontend")
busc.send_keys(FB.Keys.ENTER)


infocand = []
datex = []
FB.time.sleep(10)


for i in range (1,11):
    dir_nam="html/body/div[6]/div[3]/div[2]/div/div/main/div/div/div/ul/li["+str(i)+"]/div/div/div[2]/div/div/div/span/span/a/span/span"
    dir_ocu="html/body/div[6]/div[3]/div[2]/div/div/main/div/div/div/ul/li["+str(i)+"]/div/div/div[2]/div/div[2]/div/div"
    dir_ubi="html/body/div[6]/div[3]/div[2]/div/div/main/div/div/div/ul/li["+str(i)+"]/div/div/div[2]/div/div[2]/div/div[2]"
    dir_link="html/body/div[6]/div[3]/div[2]/div/div/main/div/div/div/ul/li["+str(i)+"]/div/div/div[2]/div/div/div/span/span/a"

    nomb = FB.D_CHR.find_element_by_xpath(dir_nam).text
    occp = FB.D_CHR.find_element_by_xpath(dir_ocu).text
    ubic = FB.D_CHR.find_element_by_xpath(dir_ubi).text
    zled = FB.D_CHR.find_element_by_xpath(dir_link).get_attribute("href")

    infocand=[nomb,occp,ubic,zled]

    datex.append(infocand)



enca = ["Nombre","ocupacion","ubicación","Link"]
datfra = FB.pd.DataFrame(datex,columns=enca)

FileName_Export = "TER.xlsx"
with FB.pd.ExcelWriter(FileName_Export, mode='w', engine='xlsxwriter') as writer:
    sheet_name = 'LinkedIn'
    datfra.to_excel(writer, sheet_name=sheet_name, index=False)
    FB.format_tbl(writer,sheet_name,datfra)

