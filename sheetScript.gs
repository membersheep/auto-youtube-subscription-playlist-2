function updatePlaylists() {
  var reservedTableRows = 3; // Row index of the first PlaylistID
  var reservedTableColumns = 2; // Column index of the first ChannelID
  var debugFlag_dontUpdatePlaylists = false;
  var flag_sendLogMail = true;

  /// VARS
  var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheets()[0];
  var data = sheet.getDataRange().getValues();
  var errorLog = '';

  /// FOR EACH PLAYLIST
  for (var currentRow = reservedTableRows; currentRow < sheet.getLastRow(); currentRow++) {
    var playlistId = data[currentRow][0];
    if (!playlistId)
      continue;

    /// GET CHANNEL IDS
    var channelIds = [];
    for (var currentColumn = reservedTableColumns; currentColumn < sheet.getLastColumn(); currentColumn++) {
      var channel = data[currentRow][currentColumn];
      if (!channel)
        continue;
      else if (channel == "ALL")
        channelIds.push.apply(channelIds, getAllChannelIds());
      else if (!(channel.substring(0,2) == "UC" && channel.length > 10)) {
        try {
          channelIds.push(YouTube.Channels.list('id', {forUsername: channel, maxResults: 1}).items[0].id);
        } catch (e) {
          Logger.log("ERROR: " + e.message + " while getting channel ids.");
          continue;
        }
      }
      else
        channelIds.push(channel);
    }

    /// GET VIDEOS FROM 24 HOURS AGO
    var videoIds = [];
    var fromDate = ISODateString(subDaysFromDate(new Date(), 1));
    for (var i = 0; i < channelIds.length; i++) {
      videoIds.push.apply(videoIds, getVideoIds(channelIds[i], fromDate));
    }
//    Logger.log("Number of vids posted from yesteday: " + videoIds.length);

    /// FILTER OUT ALREADY ADDED VIDEOS
    var playlistVideos = getVideosForPlaylist(playlistId);
    videoIds = videoIds.filter(function(videoId) {
      return playlistVideos.every(function(video) {
        return video.snippet.resourceId.videoId != videoId;
      });
    });
//    Logger.log("After filtering already added vids : " + videoIds.length);

    /// FILTER OUT WATCHED VIDEOS
    var historyPlaylistId = getHistoryPlaylistId();
    var watchedVideos = getVideosForPlaylist(historyPlaylistId);
    videoIds = videoIds.filter(function(videoId) {
      return watchedVideos.every(function(video) {
        return video.snippet.resourceId.videoId != videoId;
      });
    });
//    Logger.log("After filtering already watched vids : " + videoIds.length);

    /// REMOVE WATCHED VIDEOS FROM PLAYLIST
    var removedVideosCount = 0;
    playlistVideos.filter(function(video) {
      return watchedVideos.some(function(watchedVideo) {
        return video.snippet.resourceId.videoId == watchedVideo.snippet.resourceId.videoId;
      });
    }).forEach(function(video) {
      removedVideosCount += 1;
      removeVideoFromPlaylist(video.id);
      Utilities.sleep(1000);
    });
//    Logger.log("Number of videos removed from playlist " + videoIds.length);

    /// ADD NEW VIDEOS TO PLAYLIST
    if (!debugFlag_dontUpdatePlaylists) {
      for (var i = 0; i < videoIds.length; i++) {
        addVideoToPlaylist(videoIds[i], playlistId);
        Utilities.sleep(1000);
      }
    }

    // SEND DEBUG MAIL
    if (Logger.getLog() != null && Logger.getLog() != '' && flag_sendLogMail) {
      var recipient = Session.getActiveUser().getEmail();
      var subject = 'Youtube playlist auto-refill log';
      var body = Logger.getLog();
      MailApp.sendEmail(recipient, subject, body);
    }
  }
}

function getHistoryPlaylistId() {
  try {
    var channels = YouTube.Channels.list('contentDetails', {mine: true});
  } catch(e) {
    Logger.log("ERROR: " + e.message + " while getting history playlist id.");
  }
  return channels.items[0].contentDetails.relatedPlaylists.watchHistory;
}

function getVideoIds(channelId, fromDate) {
  var channelVideoIds = [];
  var nextPageToken = '';
  while (nextPageToken != null) {
    try {
      var channelResponse = YouTube.Search.list('id', {
        channelId: channelId,
        maxResults: 50,
        order: "date",
        publishedAfter: fromDate,
        pageToken: nextPageToken
      });
    } catch(e) {
      Logger.log("ERROR: " + e.message + " while getting recent videos from channel.");
    }
    for (var j = 0; j < channelResponse.items.length; j++) {
      var item = channelResponse.items[j];
      channelVideoIds.push(item.id.videoId)
    }
    nextPageToken = channelResponse.nextPageToken;
  }
  return channelVideoIds;
}

function addVideoToPlaylist(videoId, playlistId) {
  try {
    YouTube.PlaylistItems.insert( {
      snippet: {
        playlistId: playlistId,
        resourceId: {
          videoId: videoId,
          kind: 'youtube#video'
        }
      }
    }, 'snippet,contentDetails');
  } catch (e) {
    Logger.log("ERROR: " + e.message + " while adding video to playlist");
  }
}

function removeVideoFromPlaylist(videoId) {
  try {
    YouTube.PlaylistItems.remove(videoId);
  } catch (e) {
    Logger.log("ERROR: " + e.message + " while removing video from playlist");
  }
}

function getVideosForPlaylist(playlistId) {
  var playlistVideos = [];
  var nextPageToken = '';
  while (nextPageToken != null) {
    try {
      var playlistResponse = YouTube.PlaylistItems.list('snippet', {
      playlistId: playlistId,
      maxResults: 50,
      pageToken: nextPageToken
    });
    } catch(e) {
      Logger.log("ERROR: " + e.message + " while getting videos from playlist");
    }
    for (var j = 0; j < playlistResponse.items.length; j++) {
      var playlistItem = playlistResponse.items[j];
      playlistVideos.push(playlistItem)
    }
    nextPageToken = playlistResponse.nextPageToken;
  }
  return playlistVideos;
}

function getAllChannelIds() { // get YT Subscriptions-List, src: https://www.reddit.com/r/youtube/comments/3br98c/a_way_to_automatically_add_subscriptions_to/
  var AboResponse, AboList = [[],[]], nextPageToken = [], nptPage = 0, i, ix;

  // Workaround: nextPageToken API-Bug (this Tokens are limited to 1000 Subscriptions... but you can add more Tokens.)
  nextPageToken = ['','CDIQAA','CGQQAA','CJYBEAA','CMgBEAA','CPoBEAA','CKwCEAA','CN4CEAA','CJADEAA','CMIDEAA','CPQDEAA','CKYEEAA','CNgEEAA','CIoFEAA','CLwFEAA','CO4FEAA','CKAGEAA','CNIGEAA','CIQHEAA','CLYHEAA'];
  try {
    do {
      AboResponse = YouTube.Subscriptions.list('snippet', {
        mine: true,
        maxResults: 50,
        order: 'alphabetical',
        pageToken: nextPageToken[nptPage],
        fields: 'items(snippet(title,resourceId(channelId)))'
      });
      for (i = 0, ix = AboResponse.items.length; i < ix; i++) {
        AboList[0][AboList[0].length] = AboResponse.items[i].snippet.title;
        AboList[1][AboList[1].length] = AboResponse.items[i].snippet.resourceId.channelId;
      }
      nptPage += 1;
    } while (AboResponse.items.length > 0 && nptPage < 20);
    if (AboList[0].length !== AboList[1].length) {
      return 'Length Title != ChannelId'; // returns a string === error
    }
  } catch (e) {
    return e;
  }
  return AboList[1];
}

function subDaysFromDate(date,d){
  var result = new Date(date.getTime()-d*(24*3600*1000));
  return result
}

function ISODateString(d) { // modified from src: http://stackoverflow.com/questions/7244246/generate-an-rfc-3339-timestamp-similar-to-google-tasks-api
 function pad(n){return n<10 ? '0'+n : n}
 return d.getUTCFullYear()+'-'
      + pad(d.getUTCMonth()+1)+'-'
      + pad(d.getUTCDate())+'T'
      + pad(d.getUTCHours())+':'
      + pad(d.getUTCMinutes())+':'
      + pad(d.getUTCSeconds())+'.000Z'
}

function onOpen() {
  SpreadsheetApp.getActiveSpreadsheet().addMenu("Functions", [{name: "Update Playlists", functionName: "updatePlaylists"}]);
}
