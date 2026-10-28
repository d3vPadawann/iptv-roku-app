sub init()
    m.top.functionName = "executeRequest"
end sub

sub executeRequest()
    url = m.top.requestUrl
    if url = "" then return

    http = CreateObject("roUrlTransfer")
    http.SetUrl(url)
    http.SetCertificatesFile("common:/certs/ca-bundle.crt")
    http.InitClientCertificates()
    http.RetainBodyOnError(true)

    response = http.GetToString()
    m.top.responseJson = response
end sub