sub init()
    m.top.functionName = "executeRequest"
end sub

sub executeRequest()
    url = m.top.requestUrl
    if url = "" then return

    if m.top.taskMode = "resolve_live" then
        resolveLiveStream(url)
        return
    end if

    http = CreateObject("roUrlTransfer")
    http.SetUrl(url)
    http.SetCertificatesFile("common:/certs/ca-bundle.crt")
    http.InitClientCertificates()
    http.RetainBodyOnError(true)
    http.AddHeader("User-Agent", "IPTVSmartersPlayer")

    response = http.GetToString()
    m.top.responseJson = response
end sub

sub resolveLiveStream(url as String)
    port = CreateObject("roMessagePort")
    http = CreateObject("roUrlTransfer")
    http.SetMessagePort(port)
    http.SetUrl(url)
    http.SetCertificatesFile("common:/certs/ca-bundle.crt")
    http.InitClientCertificates()
    http.EnableCookies()
    http.RetainBodyOnError(true)
    http.AddHeader("User-Agent", "IPTVSmartersPlayer")

    result = {
        ok: false,
        finalUrl: url,
        httpCode: 0,
        reason: ""
    }

    if http.AsyncGetToString() then
        ' Aguarda até 3.5 segundos pelo manifesto .m3u8
        msg = wait(3500, port)
        if type(msg) = "roUrlEvent" then
            code = msg.GetResponseCode()
            result.httpCode = code
            body = msg.GetString()

            ' Captura a última URL de redirecionamento 301/302 (Load Balancer do Xtream)
            headersArr = msg.GetResponseHeadersArray()
            if headersArr <> invalid then
                for each hdrBlock in headersArr
                    if hdrBlock <> invalid then
                        for each k in hdrBlock
                            if LCase(k) = "location" then
                                loc = hdrBlock[k].Trim()
                                if Left(loc, 4) = "http" then
                                    result.finalUrl = loc
                                end if
                            end if
                        end for
                    end if
                end for
            end if

            if code >= 200 and code < 300 then
                if Instr(1, body, "#EXTM3U") > 0 then
                    result.ok = true
                else
                    result.reason = "Servidor retornou fluxo binário (TS bruto) em vez de playlist HLS (#EXTM3U)."
                end if
            else
                result.reason = "Erro HTTP " + code.ToStr() + " (" + msg.GetFailureReason() + ")"
            end if
        else
            ' Se deu timeout na leitura da string, o servidor abriu um fluxo .ts contínuo infinito em vez de .m3u8
            http.AsyncCancel()
            result.reason = "Servidor enviou stream TS contínuo (sem manifesto .m3u8 suportado pelo Roku)."
        end if
    else
        result.reason = "Falha ao iniciar conexão com o stream."
    end if

    m.top.resolvedStreamInfo = result
end sub