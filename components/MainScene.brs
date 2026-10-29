sub init()
    m.categoryList = m.top.findNode("categoryList")
    m.mediaGrid = m.top.findNode("mediaGrid")
    m.videoPlayer = m.top.findNode("videoPlayer")
    m.playlistStatus = m.top.findNode("playlistStatus")
    m.groupHeader = m.top.findNode("groupHeader")
    m.itemsHeader = m.top.findNode("itemsHeader")

    ' Configura certificados e User-Agent no Player de Vídeo para evitar bloqueio do servidor Xtream
    m.videoPlayer.SetCertificatesFile("common:/certs/ca-bundle.crt")
    m.videoPlayer.InitClientCertificates()

    httpAgent = CreateObject("roHttpAgent")
    httpAgent.AddHeader("User-Agent", "IPTVSmartersPlayer")
    httpAgent.EnableCookies()
    m.videoPlayer.SetHttpAgent(httpAgent)
    m.videoPlayer.SetCertificatesFile("common:/certs/ca-bundle.crt")
    m.videoPlayer.InitClientCertificates()

    ' Referências dos 6 botões da Topbar Horizontal
    m.tabBgs = [
        m.top.findNode("tabBg0"),
        m.top.findNode("tabBg1"),
        m.top.findNode("tabBg2"),
        m.top.findNode("tabBg3"),
        m.top.findNode("tabBg4"),
        m.top.findNode("tabBg5")
    ]

    m.topTabIndex = 0
    m.activeSectionIdx = 0
    m.focusArea = "topbar"

    m.currentSection = "live"
    m.currentCategories = []
    m.currentItems = []
    m.filteredItems = invalid
    m.showingSeriesEpisodes = false
    m.savedSeriesItems = []
    m.pendingMode = ""

    ' Controle de tentativas de reprodução (Fallback HLS -> TS)
    m.isPlayingLive = false
    m.playAttempt = 0
    m.currentStreamId = ""
    m.currentStreamTitle = ""

    loadCredentials()

    m.categoryList.observeField("itemSelected", "onCategorySelected")
    m.mediaGrid.observeField("itemSelected", "onMediaSelected")
    m.videoPlayer.observeField("state", "onVideoStateChange")

   m.apiTask = CreateObject("roSGNode", "ApiTask")
    m.apiTask.observeField("responseJson", "onApiResponse")
    m.apiTask.observeField("resolvedStreamInfo", "onLiveStreamResolved")

    setFocusArea("topbar")
    loadCategories("live")
end sub

sub setFocusArea(area as String)
    m.focusArea = area
    if area = "topbar" then
        m.categoryList.setFocus(false)
        m.mediaGrid.setFocus(false)
        m.top.setFocus(true)
        m.groupHeader.color = "0x94a3b8FF"
    else if area = "categories" then
        m.categoryList.setFocus(true)
        m.groupHeader.color = "0xf59e0bFF"
    else if area = "grid" then
        m.mediaGrid.setFocus(true)
        m.groupHeader.color = "0x94a3b8FF"
    else if area = "player" then
        m.videoPlayer.setFocus(true)
    end if
    updateTopBarVisuals()
end sub

sub updateTopBarVisuals()
    for i = 0 to 5
        if m.focusArea = "topbar" and i = m.topTabIndex then
            m.tabBgs[i].color = "0xf59e0bFF"
        else if i = m.activeSectionIdx then
            m.tabBgs[i].color = "0x3b82f6FF"
        else
            m.tabBgs[i].color = "0x1e293bFF"
        end if
    end for
end sub

sub activateTopBarItem()
    idx = m.topTabIndex
    if idx = 0 then
        m.activeSectionIdx = 0
        m.currentSection = "live"
        m.groupHeader.text = "GRUPOS DE CANAIS"
        loadCategories("live")
        setFocusArea("categories")
    else if idx = 1 then
        m.activeSectionIdx = 1
        m.currentSection = "vod"
        m.groupHeader.text = "CATEGORIAS DE FILMES"
        loadCategories("vod")
        setFocusArea("categories")
    else if idx = 2 then
        m.activeSectionIdx = 2
        m.currentSection = "series"
        m.groupHeader.text = "CATEGORIAS DE SÉRIES"
        loadCategories("series")
        setFocusArea("categories")
    else if idx = 3 then
        m.activeSectionIdx = 3
        m.currentSection = "docs"
        m.groupHeader.text = "DOCUMENTÁRIOS"
        loadCategories("docs")
        setFocusArea("categories")
    else if idx = 4 then
        openSearchDialog()
    else if idx = 5 then
        openEditPlaylistDialog()
    end if
    updateTopBarVisuals()
end sub

sub loadCredentials()
    sec = CreateObject("roRegistrySection", "SmartIPTVXtream")
    if sec.Exists("url") then
        m.serverUrl = sec.Read("url")
        m.username = sec.Read("username")
        m.password = sec.Read("password")
    else
        m.serverUrl = "http://gmacentral.lat"
        m.username = "Matheus4536"
        m.password = "613UHNupr"
        saveCredentials(m.serverUrl, m.username, m.password)
    end if
    m.playlistStatus.text = m.serverUrl + " (" + m.username + ")"
end sub

sub saveCredentials(url as String, user as String, pass as String)
    sec = CreateObject("roRegistrySection", "SmartIPTVXtream")
    sec.Write("url", url)
    sec.Write("username", user)
    sec.Write("password", pass)
    sec.Flush()
    m.serverUrl = url
    m.username = user
    m.password = pass
    m.playlistStatus.text = m.serverUrl + " (" + m.username + ")"
end sub

function valToStr(v as Dynamic) as String
    if v = invalid then return ""
    t = type(v)
    if t = "roString" or t = "String" then return v
    if t = "roInt" or t = "roInteger" or t = "Integer" then return v.ToStr()
    if t = "roFloat" or t = "Float" or t = "roDouble" or t = "Double" then return Str(Int(v)).Trim()
    return ""
end function

function buildApiUrl(action as String, extra as String) as String
    return m.serverUrl + "/player_api.php?username=" + m.username + "&password=" + m.password + "&action=" + action + extra
end function

sub runApi(mode as String, url as String)
    m.pendingMode = mode
    m.apiTask.control = "STOP"
    m.apiTask.taskMode = "api"
    m.apiTask.requestUrl = url
    m.apiTask.control = "RUN"
end sub

sub loadCategories(section as String)
    m.showingSeriesEpisodes = false
    m.filteredItems = invalid
    m.itemsHeader.text = "Carregando grupos..."
    if section = "live" or section = "docs" then
        runApi("categories", buildApiUrl("get_live_categories", ""))
    else if section = "vod" then
        runApi("categories", buildApiUrl("get_vod_categories", ""))
    else if section = "series" then
        runApi("categories", buildApiUrl("get_series_categories", ""))
    end if
end sub

sub onApiResponse()
    raw = m.apiTask.responseJson
    if raw = invalid or raw = "" then
        m.itemsHeader.text = "Erro ao conectar em " + m.serverUrl
        return
    end if

    parsed = ParseJson(raw)
    if parsed = invalid then
        m.itemsHeader.text = "Resposta inválida do servidor."
        return
    end if

    if m.pendingMode = "categories" then
        renderCategoryList(parsed)
    else if m.pendingMode = "streams" then
        renderMediaGrid(parsed)
    else if m.pendingMode = "series_info" then
        renderSeriesEpisodesInGrid(parsed)
    end if
end sub

sub renderCategoryList(categories as Object)
    listContent = CreateObject("roSGNode", "ContentNode")
    m.currentCategories = []

    if type(categories) = "roArray" then
        for each cat in categories
            catName = valToStr(cat.category_name)
            if m.currentSection = "docs" then
                lower = LCase(catName)
                if Instr(1, lower, "doc") > 0 or Instr(1, lower, "hist") > 0 or Instr(1, lower, "discovery") > 0 or Instr(1, lower, "nature") > 0 then
                    node = listContent.CreateChild("ContentNode")
                    node.title = catName
                    m.currentCategories.Push(cat)
                end if
            else
                node = listContent.CreateChild("ContentNode")
                node.title = catName
                m.currentCategories.Push(cat)
            end if
        end for
    end if

    m.categoryList.content = listContent
    if m.currentCategories.Count() > 0 then
        loadGroupItems(0)
    else
        m.itemsHeader.text = "Nenhum grupo encontrado."
    end if
end sub

sub onCategorySelected()
    idx = m.categoryList.itemSelected
    loadGroupItems(idx)
end sub

sub loadGroupItems(index as Integer)
    if index < 0 or index >= m.currentCategories.Count() then return
    m.showingSeriesEpisodes = false
    m.filteredItems = invalid
    cat = m.currentCategories[index]
    catId = valToStr(cat.category_id)
    catName = valToStr(cat.category_name)
    m.itemsHeader.text = catName + " (Carregando canais...)"

    action = "get_live_streams"
    if m.currentSection = "vod" then action = "get_vod_streams"
    if m.currentSection = "series" then action = "get_series"

    runApi("streams", buildApiUrl(action, "&category_id=" + catId))
end sub

sub renderMediaGrid(items as Object)
    if type(items) <> "roArray" then
        m.itemsHeader.text = "Nenhum item retornado."
        return
    end if
    m.currentItems = items
    populateGrid(items)
end sub

sub populateGrid(items as Object)
    gridContent = CreateObject("roSGNode", "ContentNode")
    count = 0

    for each item in items
        if count >= 200 then exit for
        node = gridContent.CreateChild("ContentNode")
        title = valToStr(item.name)
        if title = "" then title = valToStr(item.title)
        if title = "" then title = "Sem título"
        node.shortDescriptionLine1 = title

        icon = valToStr(item.stream_icon)
        if icon = "" then icon = valToStr(item.cover)
        if icon <> "" then
            node.hdGridPosterUrl = icon
        end if
        count = count + 1
    end for

    m.mediaGrid.content = gridContent
    m.itemsHeader.text = count.ToStr() + " itens no grupo"
end sub

function getActiveGridList() as Object
    if m.filteredItems <> invalid then return m.filteredItems
    return m.currentItems
end function

sub onMediaSelected()
    idx = m.mediaGrid.itemSelected
    activeList = getActiveGridList()
    if idx < 0 or idx >= activeList.Count() then return
    item = activeList[idx]

    if m.showingSeriesEpisodes then
        m.isPlayingLive = false
        epId = valToStr(item.id)
        ext = LCase(valToStr(item.container_extension))
        if ext = "" then ext = "mp4"
        url = m.serverUrl + "/series/" + m.username + "/" + m.password + "/" + epId + "." + ext
        launchVideoNode(url, valToStr(item.title), ext, false)
        return
    end if

    if m.currentSection = "live" or m.currentSection = "docs" then
        m.isPlayingLive = true
        m.currentStreamId = valToStr(item.stream_id)
        m.currentStreamTitle = valToStr(item.name)

        hlsUrl = m.serverUrl + "/live/" + m.username + "/" + m.password + "/" + m.currentStreamId + ".m3u8"
        m.itemsHeader.text = "Conectando ao canal " + m.currentStreamTitle + "..."

        m.apiTask.control = "STOP"
        m.apiTask.taskMode = "resolve_live"
        m.apiTask.requestUrl = hlsUrl
        m.apiTask.control = "RUN"
    else if m.currentSection = "vod" then
        m.isPlayingLive = false
        streamId = valToStr(item.stream_id)
        ext = LCase(valToStr(item.container_extension))
        if ext = "" then ext = "mp4"
        url = m.serverUrl + "/movie/" + m.username + "/" + m.password + "/" + streamId + "." + ext
        launchVideoNode(url, valToStr(item.name), ext, false)
    else if m.currentSection = "series" then
        seriesId = valToStr(item.series_id)
        m.itemsHeader.text = "Carregando episódios de " + valToStr(item.name) + "..."
        runApi("series_info", buildApiUrl("get_series_info", "&series_id=" + seriesId))
    end if
end sub

sub renderSeriesEpisodesInGrid(info as Object)
    if info = invalid or info.episodes = invalid then
        m.itemsHeader.text = "Nenhum episódio disponível."
        return
    end if

    epList = []
    for each seasonKey in info.episodes
        seasonEps = info.episodes[seasonKey]
        if type(seasonEps) = "roArray" then
            for each ep in seasonEps
                epItem = {
                    id: ep.id,
                    title: "T" + valToStr(seasonKey) + "E" + valToStr(ep.episode_num) + " - " + valToStr(ep.title),
                    name: "T" + valToStr(seasonKey) + "E" + valToStr(ep.episode_num) + " - " + valToStr(ep.title),
                    container_extension: ep.container_extension,
                    stream_icon: ""
                }
                epList.Push(epItem)
            end for
        end if
    end for

    if epList.Count() = 0 then
        m.itemsHeader.text = "Nenhum episódio encontrado."
        return
    end if

    m.savedSeriesItems = m.currentItems
    m.showingSeriesEpisodes = true
    m.currentItems = epList
    m.filteredItems = invalid
    populateGrid(epList)
    m.itemsHeader.text = epList.Count().ToStr() + " episódios (Aperte Voltar para retornar às séries)"
end sub

sub onLiveStreamResolved()
    info = m.apiTask.resolvedStreamInfo
    if info = invalid then return

    if info.ok = true then
        ' Toca diretamente a URL final do Load Balancer (evita erro de caminho relativo após 302 no Roku)
        launchVideoNode(info.finalUrl, m.currentStreamTitle, "hls", true)
    else
        m.itemsHeader.text = "Canal incompatível/indisponível: " + valToStr(info.reason)
    end if
end sub

sub launchVideoNode(url as String, title as String, format as String, isLive as Boolean)
    m.videoPlayer.control = "stop"
    m.videoPlayer.content = invalid

    videoContent = createObject("RoSGNode", "ContentNode")
    videoContent.url = url
    videoContent.title = title
    videoContent.live = isLive

    fmt = LCase(format)
    if fmt <> "hls" and fmt <> "mp4" and fmt <> "mkv" and fmt <> "ism" and fmt <> "dash" then
        fmt = "hls"
    end if
    videoContent.streamFormat = fmt

    if isLive and fmt = "hls" then
        videoContent.switchingStrategy = "full-adaptation"
    end if

    ' Não defina videoContent.HttpHeaders aqui para não sobrescrever o roHttpAgent com Cookies do VideoPlayer
    m.videoPlayer.content = videoContent
    m.videoPlayer.visible = true
    setFocusArea("player")
    m.videoPlayer.control = "play"
end sub

sub onVideoStateChange()
    state = m.videoPlayer.state

    if state = "finished" then
        stopPlayback()
    else if state = "error" then
        errCode = valToStr(m.videoPlayer.errorCode)
        errMsg = valToStr(m.videoPlayer.errorMsg)
        errStr = valToStr(m.videoPlayer.errorStr)
        m.itemsHeader.text = "Erro no Player (" + errCode + ": " + errMsg + " " + errStr + ")"
        stopPlayback()
    end if
end sub

sub stopPlayback()
    m.videoPlayer.control = "stop"
    m.videoPlayer.visible = false
    setFocusArea("grid")
end sub

sub openSearchDialog()
    kb = CreateObject("roSGNode", "KeyboardDialog")
    kb.title = "Pesquisar no Grupo Atual"
    kb.buttons = ["Buscar", "Limpar Filtro"]
    kb.observeField("buttonSelected", "onSearchDialogButton")
    m.top.dialog = kb
end sub

sub onSearchDialogButton()
    dlg = m.top.dialog
    btn = dlg.buttonSelected
    query = LCase(dlg.text)
    dlg.close = true

    if btn = 0 and query <> "" then
        filtered = []
        for each item in m.currentItems
            name = LCase(valToStr(item.name))
            if Instr(1, name, query) > 0 then filtered.Push(item)
        end for
        m.filteredItems = filtered
        populateGrid(filtered)
        setFocusArea("grid")
    else
        m.filteredItems = invalid
        populateGrid(m.currentItems)
        setFocusArea("topbar")
    end if
end sub

sub openEditPlaylistDialog()
    kb = CreateObject("roSGNode", "KeyboardDialog")
    kb.title = "Editar Lista (URL|Usuario|Senha)"
    kb.text = m.serverUrl + "|" + m.username + "|" + m.password
    kb.buttons = ["Salvar e Conectar", "Cancelar"]
    kb.observeField("buttonSelected", "onEditPlaylistButton")
    m.top.dialog = kb
end sub

sub onEditPlaylistButton()
    dlg = m.top.dialog
    if dlg.buttonSelected = 0 then
        parts = dlg.text.Split("|")
        if parts.Count() = 3 then
            saveCredentials(parts[0].Trim(), parts[1].Trim(), parts[2].Trim())
            loadCategories(m.currentSection)
        end if
    end if
    dlg.close = true
    setFocusArea("topbar")
end sub

function onKeyEvent(key as String, press as Boolean) as Boolean
    if not press then return false

    if m.focusArea = "player" then
        if key = "back" then
            stopPlayback()
            return true
        end if
        return false
    end if

    if m.focusArea = "topbar" then
        if key = "left" then
            if m.topTabIndex > 0 then
                m.topTabIndex = m.topTabIndex - 1
                updateTopBarVisuals()
            end if
            return true
        else if key = "right" then
            if m.topTabIndex < 5 then
                m.topTabIndex = m.topTabIndex + 1
                updateTopBarVisuals()
            end if
            return true
        else if key = "down" then
            setFocusArea("categories")
            return true
        else if key = "OK" then
            activateTopBarItem()
            return true
        end if
        return false
    end if

    if m.focusArea = "categories" then
        if key = "right" then
            setFocusArea("grid")
            return true
        else if key = "up" or key = "back" then
            setFocusArea("topbar")
            return true
        end if
        return false
    end if

    if m.focusArea = "grid" then
        if key = "left" then
            setFocusArea("categories")
            return true
        else if key = "up" then
            setFocusArea("topbar")
            return true
        else if key = "back" then
            if m.showingSeriesEpisodes then
                m.showingSeriesEpisodes = false
                m.currentItems = m.savedSeriesItems
                populateGrid(m.currentItems)
                return true
            end if
            setFocusArea("categories")
            return true
        end if
        return false
    end if

    return false
end function