###----------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
###  Overview: Builds the popular music playlists, then shares them with your friends
###
###  Details: Using popular music track data in plex
###  Generates top playlists for each genre where the tracks represent the #1 ranked recently released songs for artists in that genre.
###  Generates new playlists for each genre where the tracks represent the newest popular songs for artists in that genre.
###  Generates discovery playlists for each genre where the tracks represent the ranked recently released songs for artists in that genre that are not the #1 most popular song for that artist.
###
###----------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------

##Settings
$libraryTitle = "Music" ##Name of your music library
$playlistSummaryTag = "Managed Playlist" #This tag will appear in the summary of the playlist, it is use to track which playlists were created dynamically.

Add-Type -AssemblyName System.Web

function Get-Headers([String]$token){
    @{'X-Plex-Token' = $token;
      'Accept' = 'application/json';
     }
}

function Get-BaseUri($server){
    $server.scheme + "://" + $server.host + ":" + $server.port
}

function Get-Server([String]$token, [String]$name = $null){    
    $result = (Invoke-RestMethod -Uri "https://plex.tv/api/servers" -Method GET -Headers (Get-Headers -token $token)).MediaContainer.Server | Where {$_.owned -eq 1}
    if($result -eq $null){throw "No servers found."}
    if([String]::IsNullOrEmpty($name)){ $result } else { $result | Where {$_.name -eq $name} }    
}

function Get-Friends($server, [String]$username = $null){
    $result = (Invoke-RestMethod -Uri "https://plex.tv/api/servers/$($server.machineIdentifier)/access_tokens.xml?includeProfiles=1&includeProviders=1" -Method Get -Headers (Get-Headers -token $server.accessToken)).access_tokens.access_token | where {$_.owned -eq 0}
    if([String]::IsNullOrEmpty($username)){ $result } else { $result | Where {$_.username -eq $username} }
}

function Get-Identifier($server){
    (Invoke-RestMethod -Uri ((Get-BaseUri -server $server) + "/library") -Method Get -Headers (Get-Headers -token $server.accessToken)).MediaContainer.identifier
}

function Get-Libraries($server, [string]$title = $null, [string]$agent = $null){
    $result = (Invoke-RestMethod -Uri ((Get-BaseUri -server $server) + "/library/sections") -Method Get -Headers (Get-Headers -token $server.accessToken)).MediaContainer.Directory
    if(![String]::IsNullOrEmpty($title)){ $result = $result | Where {$_.title -eq $title} }
    if(![String]::IsNullOrEmpty($agent)){ $result = $result | Where {$_.agent -eq $agent} }
    $result
}

function Get-MusicGenres($server, $library){
    (Invoke-RestMethod -Uri ((Get-BaseUri -server $server) + "/library/sections/" + $library.key + "/genre") -Method Get -Headers (Get-Headers -token $server.accessToken)).MediaContainer.Directory
}

function Get-MusicArtists($server, $library, $genere = $null){
    if($genere -eq $null){
        (Invoke-RestMethod -Uri ((Get-BaseUri -server $server) + "/library/sections/" + $library.key + "/all") -Method Get -Headers (Get-Headers -token $server.accessToken)).MediaContainer.Metadata
    }
    else {
        (Invoke-RestMethod -Uri ((Get-BaseUri -server $server) + $genre.fastkey) -Method Get -Headers (Get-Headers -token $server.accessToken)).MediaContainer.Metadata
    }
}

function Get-MusicPopularTracks($server, $artists){
    $tracks = @()

    foreach ($artist in $artists){
        $artistMetadata = (Invoke-RestMethod -Uri ((Get-BaseUri -server $server) + "/library/metadata/" + $artist.ratingKey + "?includePopularLeaves=1") -Method Get -Headers (Get-Headers -token $server.accessToken)).MediaContainer.Metadata  
        $artistAlbums = (Invoke-RestMethod -Uri ((Get-BaseUri -server $server) + "/library/metadata/" + $artist.ratingKey + "/children") -Method Get -Headers (Get-Headers -token $server.accessToken)).MediaContainer.Metadata   

        if(($artistAlbums | Where {$_.rating -ne $null}).Count -ne 0){

            $artistMeasure = $artistAlbums | Where {$_.rating -ne $null} | Measure-Object rating -Average
            $artistRating = $artistMeasure.Average / $artistMeasure.Count
            
            $trackRank = 0
            foreach($track in $artistMetadata.PopularLeaves.Metadata){
                $trackRank ++

                $track | Add-Member -NotePropertyName originallyAvailableAt -NotePropertyValue $null
                $track | Add-Member -NotePropertyName albumRating -NotePropertyValue $null
                $track | Add-Member -NotePropertyName artistRating -NotePropertyValue $null
                $track | Add-Member -NotePropertyName trackRanking -NotePropertyValue $trackRank
            
                $track.artistRating = $artistRating

                $album = $artistAlbums | Where {$_.ratingKey -eq $track.parentRatingKey}
                if($album.rating -ne $null){
                    $track.albumRating = $album.rating

                    if($album.originallyAvailableAt -ne $null){
                        $track.originallyAvailableAt = [datetime]$album.originallyAvailableAt
                    }
                    $tracks += $track
                }    
            }
        }
    }

    $tracks
}

function Get-Playlists($server, [String]$playlistType = $null, [String]$summary = $null, [String]$title = $null ){
    $result = (Invoke-RestMethod -Uri ((Get-BaseUri -server $server) + "/playlists") -Method Get -Headers (Get-Headers -token $server.accessToken)).MediaContainer.Metadata
    if(![String]::IsNullOrEmpty($playlistType)) {$result = $result | Where {$_.playlistType -eq $playlistType}}
    if(![String]::IsNullOrEmpty($summary)) {$result = $result | Where {$_.summary -eq $summary}}
    if(![String]::IsNullOrEmpty($title)) {$result = $result | Where {$_.title -eq $title}}
    $result
}

function Get-MusicPlaylistItems($server, $playlist){
    (Invoke-RestMethod -Uri ((Get-BaseUri -server $server) + $playlist.key) -Method Get -Headers (Get-Headers -token $server.accessToken)).MediaContainer.Metadata
}

function Set-MusicPlaylist($server, $identifier, [string]$title, [string]$ratingKeysString, [String]$summary){
    if(![String]::IsNullOrEmpty($ratingKeysString)){
        $uriRequest = $null
        $playlist = Get-Playlists -server $server -playlistType "audio" -summary "Managed Playlist" -title $title
        if($playlist -eq $null){

            #create playlist
            $uriRequest = (Get-BaseUri -server $server) + "/playlists"
            $uriRequest += "?uri=server://" + $server.machineIdentifier + "/" + $identifier + "/library/metadata/" + $ratingKeysString
            $uriRequest += "&includeExternalMedia=1&title=" + [System.Web.HttpUtility]::UrlEncode($title) +"&smart=0" + "&type=audio"
            $request = Invoke-RestMethod -Uri $uriRequest -Method Post -Headers (Get-Headers -token $server.accessToken)
            $request2 = Invoke-RestMethod -Uri ((Get-BaseUri -server $server) + "/playlists/" + $request.MediaContainer.Metadata[0].ratingKey + "?includeExternalMedia=1&summary=" + [System.Web.HttpUtility]::UrlEncode($summary)) -Method Put -Headers (Get-Headers -token $server.accessToken)
        }
        else{

            #update playlist
            $uriRequest = (Get-BaseUri -server $server) + $playlist.key
            $uriRequest += "?uri=server://" + $server.machineIdentifier + "/" + $identifier + "/library/metadata/" + $ratingKeysString
            $uriRequest += "&includeExternalMedia=1"
            $request = Invoke-RestMethod -Uri $uriRequest -Method Put -Headers (Get-Headers -token $server.accessToken)
        }
    }
}

function Delete-Playlist($server, $playlist){
    if($playlist -ne $null){
        $result = Invoke-RestMethod ((Get-BaseUri -server $server) + "/playlists/" + $playlist.ratingKey) -Method Delete -Headers (Get-Headers -token $server.accessToken)
    }
}

function Sync-MusicPlaylists($server, $identifier, $playlist, $friend){
    #Get items in playlist
    $ratingKeysString = (Get-MusicPlaylistItems -server $server -playlist $playlist).ratingKey -join ','
        
    if(![String]::IsNullOrEmpty($ratingKeysString)){

        ##Delete Music Playlist if it exists
        $playlistDest = (Invoke-RestMethod -Uri ((Get-BaseUri -server $server) + "/playlists") -Method Get -Headers (Get-Headers -token $friend.token)).MediaContainer.Metadata | Where {($_.title -eq $playlist.title) -and ($_.summary -eq $playlist.summary)}
        if($playlistDest -ne $null){
            $result = Invoke-RestMethod ((Get-BaseUri -server $server) + "/playlists/" + $playlistDest.ratingKey) -Method Delete -Headers (Get-Headers -token $friend.token)
        }

        ##Re/create Music Playlist        
        $uriRequest = $null
        $uriRequest = (Get-BaseUri -server $server) + "/playlists"
        $uriRequest += "?uri=server://" + $server.machineIdentifier + "/" + $identifier + "/library/metadata/" + $ratingKeysString
        $uriRequest += "&includeExternalMedia=1&title=" + [System.Web.HttpUtility]::UrlEncode($playlist.title) +"&smart=0" + "&type=audio"
        $request = Invoke-RestMethod -Uri $uriRequest -Method Post -Headers (Get-Headers -token $friend.token)
        $request2 = Invoke-RestMethod -Uri ((Get-BaseUri -server $server) + "/playlists/" + $request.MediaContainer.Metadata[0].ratingKey + "?includeExternalMedia=1&summary=" + [System.Web.HttpUtility]::UrlEncode($playlist.summary)) -Method Put -Headers (Get-Headers -token $friend.token)

    }
}

function Get-PlexToken(){
    $credential = Get-Credential -Message "Plex credentials"
    $session = Invoke-RestMethod -Uri  "https://plex.tv/users/sign_in.json" -Method Post -Credential $credential -Headers @{
            'X-Plex-Client-Identifier'="PowerShell";
            'X-Plex-Product'='PowerShell';
            'X-Plex-Version'="V0.01";
            'X-Plex-Username'=$credential.GetNetworkCredential().UserName;
            'Accept' = 'application/json';
		}
    $session.user.authentication_token
}

Write-Host ("Plex: Operation Started")
Write-Host ("Plex: Authentication")
$token = Get-PlexToken

Write-Host ("Plex: Connecting...")
Write-Host ("Plex: Getting Server")
$server = Get-Server -token $token
Write-Host ("Plex: Getting Friends")
$friends = Get-Friends -server $server
Write-Host ("Plex/$($server.name): Connecting...")
Write-Host ("Plex/$($server.name): Getting library identifier")
$identifier = Get-Identifier -server $server
Write-Host ("Plex/$($server.name)/Library/$($libraryTitle): Getting Library")
$library = Get-Libraries -server $server -agent "tv.plex.agents.music" -title $libraryTitle

Write-Host ("Plex/$($server.name)/Library/$($libraryTitle): Getting Genres")
foreach ($genre in Get-MusicGenres -server $server -library $library){
    Write-Host ("")
    Write-Host ("Plex/$($server.name)/Library/$($libraryTitle)/Genre/$($genre.title): Getting Popular Music Tracks")

    $tracks = Get-MusicPopularTracks -server $server -artists (Get-MusicArtists -server $server -library $library -genere $genre)
    $tracksTop = $tracks | Where {($_.trackRanking -eq 1) -and ($_.originallyAvailableAt -gt (Get-Date).AddYears(-4)) -and (($_.artistRating -ge 5) -or ($_.albumRating -ge 5))} | Sort-Object -Property artistRating, albumRating -Descending | Select -First 100
    $tracksNew = $tracks | Where {(($_.artistRating -ge 5) -or ($_.albumRating -ge 5))} | Sort-Object -Property originallyAvailableAt -Descending | Sort-Object -Property trackRanking | Select -First 100
    
    $tracksDiscover = @()
    if(($tracks | Where {($_.trackRanking -ne 1) -and ($_.originallyAvailableAt -gt (Get-Date).AddYears(-1)) -and (($_.artistRating -ge 7) -or ($_.albumRating -ge 7))}).Count -gt 20){
        $tracksDiscover = $tracks | Where {($_.trackRanking -ne 1) -and ($_.originallyAvailableAt -gt (Get-Date).AddYears(-1)) -and (($_.artistRating -ge 7) -or ($_.albumRating -ge 7))} | Sort-Object -Property artistRating, albumRating -Descending | Sort-Object -Property trackRanking | Select -First 100
    }
    elseif(($tracks | Where {($_.trackRanking -ne 1) -and ($_.originallyAvailableAt -gt (Get-Date).AddYears(-2)) -and (($_.artistRating -ge 7) -or ($_.albumRating -ge 7))}).Count -gt 20){
        $tracksDiscover = $tracks | Where {($_.trackRanking -ne 1) -and ($_.originallyAvailableAt -gt (Get-Date).AddYears(-2)) -and (($_.artistRating -ge 7) -or ($_.albumRating -ge 7))} | Sort-Object -Property artistRating, albumRating -Descending | Sort-Object -Property trackRanking | Select -First 100
    }
    elseif(($tracks | Where {($_.trackRanking -ne 1) -and ($_.originallyAvailableAt -gt (Get-Date).AddYears(-5)) -and (($_.artistRating -ge 7) -or ($_.albumRating -ge 7))}).Count -gt 20){
        $tracksDiscover = $tracks | Where {($_.trackRanking -ne 1) -and ($_.originallyAvailableAt -gt (Get-Date).AddYears(-5)) -and (($_.artistRating -ge 7) -or ($_.albumRating -ge 7))} | Sort-Object -Property artistRating, albumRating -Descending | Sort-Object -Property trackRanking | Select -First 100
    }
    elseif(($tracks | Where {($_.trackRanking -ne 1) -and ($_.originallyAvailableAt -gt (Get-Date).AddYears(-10)) -and (($_.artistRating -ge 7) -or ($_.albumRating -ge 7))}).Count -gt 20){
        $tracksDiscover = $tracks | Where {($_.trackRanking -ne 1) -and ($_.originallyAvailableAt -gt (Get-Date).AddYears(-10)) -and (($_.artistRating -ge 7) -or ($_.albumRating -ge 7))} | Sort-Object -Property artistRating, albumRating -Descending | Sort-Object -Property trackRanking | Select -First 100
    }
    else{
        $tracksDiscover = $tracks | Where {($_.trackRanking -ne 1) -and (($_.artistRating -ge 5) -or ($_.albumRating -ge 5))} | Sort-Object -Property artistRating, albumRating -Descending | Sort-Object -Property trackRanking | Select -First 100
    }
    
    ##Rebuilding Local Playlists
    Write-Host ("Plex/$($server.name)/Playlists/$("Top " + $genre.title): Rebuilding")
    Delete-Playlist -server $server -playlist (Get-Playlists -server $server -playlistType "audio" -summary $playlistSummaryTag -title ("Top " + $genre.title))
    Set-MusicPlaylist -server $server -identifier $identifier -title ("Top " + $genre.title) -ratingKeysString ($tracksTop.ratingKey -join ',') -summary $playlistSummaryTag

    Write-Host ("Plex/$($server.name)/Playlists/$("Discover " + $genre.title): Rebuilding")
    Delete-Playlist -server $server -playlist (Get-Playlists -server $server -playlistType "audio" -summary $playlistSummaryTag -title ("Discover " + $genre.title))    
    Set-MusicPlaylist -server $server -identifier $identifier -title ("Discover " + $genre.title) -ratingKeysString ($tracksDiscover.ratingKey -join ',') -summary $playlistSummaryTag
    
    Write-Host ("Plex/$($server.name)/Playlists/$("New " + $genre.title): Rebuilding")
    Delete-Playlist -server $server -playlist (Get-Playlists -server $server -playlistType "audio" -summary $playlistSummaryTag -title ("New " + $genre.title))
    Set-MusicPlaylist -server $server -identifier $identifier -title ("New " + $genre.title) -ratingKeysString ($tracksNew.ratingKey -join ',') -summary $playlistSummaryTag
}

 ##Sync Playlists
Write-Host ("Plex/$($server.name)/Playlist: Getting Playlists")
$playlists = Get-Playlists -server $server -playlistType "audio" -summary $playlistSummaryTag
foreach($friend in $friends){
    foreach($playlist in $playlists){
        Write-Host ("Plex/$($server.name)/Playlist/$($playlist.title): Syncing -> Plex/Friend/$($friend.username)")
        Sync-MusicPlaylists -server $server -identifier $identifier -playlist $playlist -friend $friend
    }    
}

Write-Host("Plex: Operation Completed")
