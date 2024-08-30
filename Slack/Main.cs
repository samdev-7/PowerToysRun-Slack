using ManagedCommon;
using Microsoft.PowerToys.Settings.UI.Library;
using System;
using System.Collections;
using System.Collections.Generic;
using System.Collections.Specialized;
using System.Globalization;
using System.Linq;
using System.Net.Http;
using System.Text;
using System.Web;
using System.Windows.Automation;
using System.Windows.Controls;
using Wox.Infrastructure;
using Wox.Plugin;
using Wox.Plugin.Logger;
using System.Net.Http.Json;
using System.Net.Http.Headers;
using System.IO;
using System.Text.Json;
using System.Text.Json.Nodes;
using System.Diagnostics;
using Wox.Plugin.Common;
using System.Security.Policy;
using System.Threading.Channels;
using System.Xml.Linq;

enum ChannelType
{
    Public,
    Private,
    Mpim,    // Group DM
    Im       // DM
}

struct Channel
{
    public string id;
    public string name;
    public ChannelType type;
    public bool isArchived;
    public string description;
    public string context_team_id;
}

namespace Community.PowerToys.Run.Plugin.Slack
{
    public class Main : IPlugin, IPluginI18n, IContextMenu, ISettingProvider, IReloadable, IDisposable, IDelayedExecutionPlugin
    {
        private PluginInitContext _context;

        private string _iconPath;
        private string _channelIconPath;
        private string _privateIconPath;
        private string _archiveIconPath;

        private bool _disposed;

        public string Name => Properties.Resources.plugin_name;

        public string Description => Properties.Resources.plugin_description;

        public static string PluginID => "72564bae0b4d45408961ca91ab50a293";

        private static readonly HttpClient client = new();

        // Settings
        private string _slackToken = string.Empty;
        private bool _slackShowArchived = false;
        private bool _slackIncludePublic = true;
        private bool _slackIncludePrivate = true;
        private bool _slackIncludeMpim = true;
        private bool _slackIncludeIm = true;
        private string _slackTeam = string.Empty;

        // TODO: add additional options (optional)
        public IEnumerable<PluginAdditionalOption> AdditionalOptions =>
        [
            new()
            {
                Key = "SlackToken",
                DisplayLabel = "Slack token",
                DisplayDescription="A Bot Token or User Token with scopes channels:read, groups:read, im:read, and mpim:read",
                PluginOptionType = PluginAdditionalOption.AdditionalOptionType.Textbox,
                TextValue = string.Empty,
            },
            new()
            {
                Key="SlackArchive",
                DisplayLabel="Include archived channels",
                PluginOptionType = PluginAdditionalOption.AdditionalOptionType.Checkbox,
                Value = false
            },
            new()
            {
                Key = "SlackTypesPublic",
                DisplayLabel = "Include public channels",
                DisplayDescription="This option will be ignored with searches starting with \"#\"",
                PluginOptionType = PluginAdditionalOption.AdditionalOptionType.Checkbox,
                Value = true,
            },
            new()
            {
                Key = "SlackTypesPrivate",
                DisplayLabel = "Include private channels",
                PluginOptionType = PluginAdditionalOption.AdditionalOptionType.Checkbox,
                Value = true,
            },
            new()
            {
                Key = "SlackTypesMpim",
                DisplayLabel = "Include group DMs",
                DisplayDescription = "This option will be ignored with searches starting with \"^\"",
                PluginOptionType = PluginAdditionalOption.AdditionalOptionType.Checkbox,
                Value = false,
            },
            new()
            {
                Key = "SlackTypesIm",
                DisplayLabel = "Include DMs",
                DisplayDescription = "This option will be ignored with searches starting with \"@\"",
                PluginOptionType = PluginAdditionalOption.AdditionalOptionType.Checkbox,
                Value = false,
            },
            new()
            {
                Key = "SlackTeam",
                DisplayLabel = "Team ID",
                DisplayDescription="Leave empty for all teams",
                PluginOptionType = PluginAdditionalOption.AdditionalOptionType.Textbox,
                TextValue = string.Empty,
            }
        ];

        public void UpdateSettings(PowerLauncherPluginSettings settings)
        {
            _slackToken = settings?.AdditionalOptions?.FirstOrDefault(x => x.Key == "SlackToken")?.TextValue ?? string.Empty;
            _slackShowArchived = settings?.AdditionalOptions?.FirstOrDefault(x => x.Key == "SlackArchive")?.Value ?? false;
            _slackIncludePublic = settings?.AdditionalOptions?.FirstOrDefault(x => x.Key == "SlackTypesPublic")?.Value ?? true;
            _slackIncludePrivate = settings?.AdditionalOptions?.FirstOrDefault(x => x.Key == "SlackTypesPrivate")?.Value ?? true;
            _slackIncludeMpim = settings?.AdditionalOptions?.FirstOrDefault(x => x.Key == "SlackTypesMpim")?.Value ?? true;
            _slackIncludeIm = settings?.AdditionalOptions?.FirstOrDefault(x => x.Key == "SlackTypesIm")?.Value ?? true;
            _slackTeam = settings?.AdditionalOptions?.FirstOrDefault(x => x.Key == "SlackTeam")?.TextValue ?? string.Empty;
        }

        private static string ConstructTypeQuery(bool incPublic, bool incPrivate, bool incMpim, bool incIm)
        {
            List<string> types = [];
            if (incPublic)
            {
                types.Add("public_channel");
            }
            if (incPrivate)
            {
                types.Add("private_channel");
            }
            if (incMpim)
            {
                types.Add("mpim");
            }
            if (incIm)
            {
                types.Add("im");
            }

            return string.Join(",", types);
        }

        private async Task<List<Channel>> FetchChannels(string searchQuery)
        {
            string cursor = string.Empty;
            List<Channel> channels = [];

            bool incPublic = _slackIncludePublic;
            bool incPrivate = _slackIncludePrivate;
            bool incMpim = _slackIncludeMpim;
            bool incIm = _slackIncludeIm;

            searchQuery = searchQuery.Trim();

            if (searchQuery.StartsWith('#'))
            {
                searchQuery = searchQuery[1..];
                incPublic = true;
                incMpim = false;
                incIm = false;
            }
            if (searchQuery.StartsWith('^'))
            {
                searchQuery = searchQuery[1..];
                incPublic = false;
                incPrivate = false;
                incMpim = true;
                incIm = false;
            }
            if (searchQuery.StartsWith('@'))
            {
                searchQuery = searchQuery[1..];
                incPublic = false;
                incPrivate = false;
                incMpim = false;
                incIm = true;
            }

            if (string.IsNullOrEmpty(searchQuery))
            {
                return channels;
            }

            string types = ConstructTypeQuery(incPublic, incPrivate, incMpim, incIm);

            do
            {
                NameValueCollection query = HttpUtility.ParseQueryString(string.Empty);
                if (!string.IsNullOrEmpty(cursor))
                {
                    query["cursor"] = cursor;
                }
                query["exclude_archived"] = (!_slackShowArchived).ToString();
                query["limit"] = "1000";
                query["types"] = types;
                if (!string.IsNullOrEmpty(_slackTeam))
                {
                    query["team_id"] = _slackTeam;
                }

                HttpRequestMessage request = new()
                {
                    RequestUri = new Uri($"https://slack.com/api/users.conversations?{query}"),
                    Method = HttpMethod.Get,
                };
                request.Headers.Authorization = new AuthenticationHeaderValue("Bearer", _slackToken);

                HttpResponseMessage response = await client.SendAsync(request);
                response.EnsureSuccessStatusCode();

                if (response.Content is null)
                {
                    return channels;
                }
                if (response.Content?.Headers?.ContentType?.MediaType != "application/json")
                {
                    return channels;    
                }

                JsonNode? json = await response.Content.ReadFromJsonAsync<JsonNode>();

                if (json is null)
                {
                    return channels;
                }
                if ((bool?)json["ok"] != true)
                {
                    return channels;
                }

                if (json["channels"] is null)
                {
                    return channels;
                }

                JsonNode? jsonChannels = json["channels"];
                if (jsonChannels is null)
                {
                    return channels;
                }

                foreach (dynamic? channel in jsonChannels.AsArray())
                {
                    if (channel is null)
                    {
                        continue;
                    }
                    ChannelType type = ChannelType.Public;
                    if ((bool?)channel["is_im"] == true)
                    {
                        type = ChannelType.Im;
                    }
                    else if ((bool?)channel["is_mpim"] == true)
                    {
                        type = ChannelType.Mpim;
                    }
                    else if ((bool?)channel["is_private"] == true)
                    {
                        type = ChannelType.Private;
                    }

                    string? description = (string?)channel["purpose"]?["value"];
                    if (string.IsNullOrEmpty(description))
                    {
                        description = (string?)channel["topic"]?["value"];
                    }

                    channels.Add(new Channel
                    {
                        id = (string)channel["id"],
                        name = (string)channel["name"],
                        type = type,
                        isArchived = (bool)channel["is_archived"],
                        description = description ?? string.Empty,
                        context_team_id = (string)channel["context_team_id"],
                    });
                }

                cursor = (string?)json["response_metadata"]?["next_cursor"] ?? string.Empty;
            } while (!string.IsNullOrEmpty(cursor));

            return channels;
        }

        private static List<Channel> SearchChannels(List<Channel> channels, string query)
        {
            List<string> names = [];

            if (query.StartsWith('#'))
            {
                query = query[1..];
            }
            if (query.StartsWith('^'))
            {
                query = query[1..];
            }
            if (query.StartsWith('@'))
            {
                query = query[1..];
            }

            foreach (Channel channel in channels)
            {
                if (channel.type != ChannelType.Im && channel.type != ChannelType.Mpim)
                {
                    names.Add(channel.name);
                }
            }

            var result =  FuzzySharp.Process.ExtractTop(query, names, limit: 5);

            List<Channel> matches = [];
            foreach (var match in result)
            {
                matches.Add(channels[match.Index]);
            }

            return matches;
        }

        // TODO: return context menus for each Result (optional)
        public List<ContextMenuResult> LoadContextMenus(Result selectedResult)
        {
            return new List<ContextMenuResult>(0);
        }

        // TODO: return query results
        public List<Result> Query(Query query)
        {
            ArgumentNullException.ThrowIfNull(query);

            List<Result> results = [];

            if (!string.IsNullOrEmpty(query.Search))
            {
                List<Channel> channels = FetchChannels(query.Search).Result;
                List<Channel> matches = SearchChannels(channels, query.Search);

                foreach(Channel channel in matches)
                {
                    string subTitle = channel.description;
                    if (string.IsNullOrEmpty(subTitle))
                    {
                        subTitle = string.Empty;
                    }
                    string name = channel.name;
                    if (channel.isArchived)
                    {
                        name += " (archived)";
                    }

                    string icon = _iconPath;

                    if (channel.isArchived)
                    {
                        icon = _archiveIconPath;
                    }
                    else if (channel.type == ChannelType.Public)
                    {
                        icon = _channelIconPath;
                    }
                    else if (channel.type == ChannelType.Private)
                    {
                        icon = _privateIconPath;
                    }

                    results.Add(new Result
                    {
                        Title = name,
                        SubTitle = subTitle,
                        QueryTextDisplay = name,
                        IcoPath = icon,
                        Action = action =>
                        {
                            ProcessStartInfo startInfo = new()
                            {
                                FileName = "cmd.exe",
                                Arguments = $"/c start \"Starting Slack\" \"slack://channel?team={channel.context_team_id}&id={channel.id}\"",
                                CreateNoWindow = true,
                                UseShellExecute = false
                            };
                            Process.Start(startInfo );
                            return true;
                        },
                    });
                }
                return results;
            } else
            {
                results.Add(new Result
                {
                    Title = "Slack",
                    SubTitle = "Open Slack",
                    IcoPath = _iconPath,
                    Action = action =>
                    {
                        ProcessStartInfo startInfo = new()
                        {
                            FileName = "cmd.exe",
                            Arguments = $"/c start \"Starting Slack\" \"slack://\"",
                            CreateNoWindow = true,
                            UseShellExecute = false
                        };
                        Process.Start(startInfo);
                        return true;
                    },
                });
            }

            return results;
        }

        // TODO: return delayed query results (optional)
        public List<Result> Query(Query query, bool delayedExecution)
        {
            ArgumentNullException.ThrowIfNull(query);

            var results = new List<Result>();

            // empty query
            if (string.IsNullOrEmpty(query.Search))
            {
                return results;
            }

            return results;
        }

        public void Init(PluginInitContext context)
        {
            _context = context ?? throw new ArgumentNullException(nameof(context));
            _context.API.ThemeChanged += OnThemeChanged;
            UpdateIconPath(_context.API.GetCurrentTheme());
        }

        public string GetTranslatedPluginTitle()
        {
            return Properties.Resources.plugin_name;
        }

        public string GetTranslatedPluginDescription()
        {
            return Properties.Resources.plugin_description;
        }

        private void OnThemeChanged(Theme oldtheme, Theme newTheme)
        {
            UpdateIconPath(newTheme);
        }

        private void UpdateIconPath(Theme theme)
        {
            if (theme == Theme.Light || theme == Theme.HighContrastWhite)
            {
                _iconPath = "Images/Slack.png";
                _channelIconPath = "Images/Channel.light.png";
                _privateIconPath = "Images/Private.light.png";
                _archiveIconPath = "Images/Archive.light.png";
            }
            else
            {
                _iconPath = "Images/Slack.png";
                _channelIconPath = "Images/Channel.dark.png";
                _privateIconPath = "Images/Private.dark.png";
                _archiveIconPath = "Images/Archive.dark.png";
            }
        }

        public Control CreateSettingPanel()
        {
            throw new NotImplementedException();
        }

        public void ReloadData()
        {
            if (_context is null)
            {
                return;
            }

            UpdateIconPath(_context.API.GetCurrentTheme());
        }

        public void Dispose()
        {
            Dispose(true);
            GC.SuppressFinalize(this);
        }

        protected virtual void Dispose(bool disposing)
        {
            if (!_disposed && disposing)
            {
                if (_context != null && _context.API != null)
                {
                    _context.API.ThemeChanged -= OnThemeChanged;
                }

                _disposed = true;
            }
        }
    }
}
