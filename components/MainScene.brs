sub init()
    m.sectionTabs = m.top.findNode("sectionTabs")
    m.categoryList = m.top.findNode("categoryList")
    m.mediaGrid = m.top.findNode("mediaGrid")
    m.videoPlayer = m.top.findNode("videoPlayer")
    m.playlistStatus = m.top.findNode("playlistStatus")
    m.groupHeader = m.top.findNode("groupHeader")
    m.itemsHeader = m.top.findNode("itemsHeader")

    ' Carrega ou inicializa credenciais salvas no Registro da TV Roku
    loadCredentials()

    ' Configura Abas da Topbar
    tabsContent = CreateObject("roSGNode", "ContentNode")
    addTab(tabsContent, "Canais ao Vivo")
    addTab(tabsContent, "Filmes")
    addTab(tabsContent, "Séries")
    addTab(tabsContent, "Documentários")
    addTab(tabsContent, "Pesquisar")
    addTab(tabsContent, "Configurar Lista")
    m.sectionTabs.content = tabsContent

    m.currentSection = "live"
    m.currentCategories = []
    m.currentItems = []
    m.pendingMode = ""

    m.sectionTabs.observeField("itemSelected", "onTabSelected")
    m.categoryList.observeField("itemSelected", "onCategorySelected")
    m.mediaGrid.observeField("itemSelected", "onMediaSelected")
    m.videoPlayer.observeField("state", "onVideoStateChange")

    m.apiTask = CreateObject("roSGNode", "ApiTask")
    m.apiTask.observeField("responseJson", "onApiResponse")

    m.sectionTabs.setFocus(true)
    loadCategories("live")
end sub

sub addTab(parent as Object, title as String)
    node = parent.CreateChild("ContentNode")
    node.title = title
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

function buildApiUrl(action as String, extra as String) as String
    return m.serverUrl + "/player_api.php?username=" + m.username + "&password=" + m.password + "&action=" + action + extra
end function

sub runApi(mode as String, url as String)
    m.pendingMode = mode
    m.apiTask.control = "STOP"
    m.apiTask.requestUrl = url
    m.apiTask.control = "RUN"
end sub

sub onTabSelected()
    idx = m.sectionTabs.itemSelected
    if idx = 0 then
        m.currentSection = "live"
        m.groupHeader.text = "GRUPOS DE CANAIS"
        loadCategories("live")
    else if idx = 1 then
        m.currentSection = "vod"
        m.groupHeader.text = "CATEGORIAS DE FILMES"
        loadCategories("vod")
    else if idx = 2 then
        m.currentSection = "series"
        m.groupHeader.text = "CATEGORIAS DE SÉRIES"
        loadCategories("series")
    else if idx = 3 then
        m.currentSection = "docs"
        m.groupHeader.text = "DOCUMENTÁRIOS"
        loadCategories("docs")
    else if idx = 4 then
        openSearchDialog()
    else if idx = 5 then
        openEditPlaylistDialog()
    end if
end sub

sub loadCategories(section as String)
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
        playFirstSeriesEpisode(parsed)
    end if
end sub

sub renderCategoryList(categories as Object)
    listContent = CreateObject("roSGNode", "ContentNode")
    m.currentCategories = []

    for each cat in categories
        catName = cat.category_name
        if m.currentSection = "docs" then
            lower = LCase(catName)
            if Instr(1, lower, "doc") > 0 or Instr(1, lower, "hist") > 0 or Instr(1, lower, "discovery") > 0 then
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
    cat = m.currentCategories[index]
    catId = cat.category_id.ToStr()
    m.itemsHeader.text = cat.category_name + " (Carregando...)"

    action = "get_live_streams"
    if m.currentSection = "vod" then action = "get_vod_streams"
    if m.currentSection = "series" then action = "get_series"

    runApi("streams", buildApiUrl(action, "&category_id=" + catId))
end sub

sub renderMediaGrid(items as Object)
    m.currentItems = items
    populateGrid(items)
end sub

sub populateGrid(items as Object)
    gridContent = CreateObject("roSGNode", "ContentNode")
    count = 0

    for each item in items
        if count >= 200 then exit for
        node = gridContent.CreateChild("ContentNode")
        title = item.name
        if title = invalid then title = item.title
        if title = invalid then title = "Sem título"
        node.shortDescriptionLine1 = title

        icon = item.stream_icon
        if icon = invalid then icon = item.cover
        if icon <> invalid and icon <> "" then
            node.hdGridPosterUrl = icon
        end if
        count = count + 1
    end for

    m.mediaGrid.content = gridContent
    m.itemsHeader.text = count.ToStr() + " itens carregados"
end sub

sub onMediaSelected()
    idx = m.mediaGrid.itemSelected
    if idx < 0 or idx >= m.filteredOrCurrentItems().Count() then return
    item = m.filteredOrCurrentItems()[idx]

    if m.currentSection = "live" or m.currentSection = "docs" then
        streamId = item.stream_id.ToStr()
        url = m.serverUrl + "/live/" + m.username + "/" + m.password + "/" + streamId + ".m3u8"
        startPlayback(url, item.name, "hls")
    else if m.currentSection = "vod" then
        streamId = item.stream_id.ToStr()
        ext = "mp4"
        if item.container_extension <> invalid then ext = item.container_extension
        url = m.serverUrl + "/movie/" + m.username + "/" + m.password + "/" + streamId + "." + ext
        startPlayback(url, item.name, ext)
    else if m.currentSection = "series" then
        seriesId = item.series_id.ToStr()
        runApi("series_info", buildApiUrl("get_series_info", "&series_id=" + seriesId))
    end if
end sub

function filteredOrCurrentItems() as Object
    if m.filteredItems <> invalid then return m.filteredItems
    return m.currentItems
end function

sub playFirstSeriesEpisode(info as Object)
    if info = invalid or info.episodes = invalid then return
    for each seasonKey in info.episodes
        eps = info.episodes[seasonKey]
        if eps.Count() > 0 then
            ep = eps[0]
            ext = "mp4"
            if ep.container_extension <> invalid then ext = ep.container_extension
            url = m.serverUrl + "/series/" + m.username + "/" + m.password + "/" + ep.id.ToStr() + "." + ext
            startPlayback(url, ep.title, ext)
            return
        end if
    end for
end sub

sub startPlayback(url as String, title as String, format as String)
    videoContent = createObject("RoSGNode", "ContentNode")
    videoContent.url = url
    videoContent.title = title
    videoContent.streamFormat = format

    m.videoPlayer.content = videoContent
    m.videoPlayer.visible = true
    m.videoPlayer.setFocus(true)
    m.videoPlayer.control = "play"
end sub

sub onVideoStateChange()
    if m.videoPlayer.state = "finished" or m.videoPlayer.state = "error" then
        stopPlayback()
    end if
end sub

sub stopPlayback()
    m.videoPlayer.control = "stop"
    m.videoPlayer.visible = false
    m.mediaGrid.setFocus(true)
end sub

' Teclado de Busca sob demanda (só abre ao clicar em Pesquisar)
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
            name = ""
            if item.name <> invalid then name = LCase(item.name)
            if Instr(1, name, query) > 0 then filtered.Push(item)
        end for
        m.filteredItems = filtered
        populateGrid(filtered)
    else
        m.filteredItems = invalid
        populateGrid(m.currentItems)
    end if
end sub

' Edição de URL, Usuário e Senha da Lista Xtream Codes
sub openEditPlaylistDialog()
    kb = CreateObject("roSGNode", "KeyboardDialog")
    kb.title = "Editar Lista (Formato: URL|Usuario|Senha)"
    kb.text = m.serverUrl + "|" + m.username + "|" + m.password
    kb.buttons = ["Salvar e Recarregar", "Cancelar"]
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
end sub

' Navegação pelas setas Esquerda/Direita/Cima/Baixo e botão Voltar do Controle Roku
function onKeyEvent(key as String, press as Boolean) as Boolean
    if not press then return false

    if key = "back" then
        if m.videoPlayer.visible then
            stopPlayback()
            return true
        else if m.mediaGrid.hasFocus() then
            m.categoryList.setFocus(true)
            return true
        else if m.categoryList.hasFocus() then
            m.sectionTabs.setFocus(true)
            return true
        end if
    else if key = "down" then
        if m.sectionTabs.hasFocus() then
            m.categoryList.setFocus(true)
            return true
        end if
    else if key = "up" then
        if m.categoryList.hasFocus() or m.mediaGrid.hasFocus() then
            m.sectionTabs.setFocus(true)
            return true
        end if
    else if key = "right" then
        if m.categoryList.hasFocus() then
            m.mediaGrid.setFocus(true)
            return true
        end if
    else if key = "left" then
        if m.mediaGrid.hasFocus() then
            m.categoryList.setFocus(true)
            return true
        end if
    end if

    return false
end function