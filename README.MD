# Cold Email Automation System

Google Apps Script backend for cold email outreach automation.
Supports 100-300 simultaneous Gmail sender accounts.

## Architecture

**Master Control** (one spreadsheet):
- `distributor.gs` — lead assignment and distribution engine
- `replyScanner.gs` — parallel Gmail reply detection across all senders
- `bounceScanner.gs` — bounce detection via mailer-daemon scanning
- `resetDaily.gs` — nightly counter reset and inbox cleanup
- `campaignSync.gs` — campaign status cascade to masterLeads
- `webApp.gs` — 34+ HTTP POST API endpoints
- `utils.gs` — shared helper functions

**Sender** (one spreadsheet per Gmail account, up to 300):
- `sender.gs` — cold email chained trigger system
- `followup.gs` — followup email chained trigger system  
- `webApp.gs` — sender-side HTTP POST endpoints
- `utils.gs` — shared helpers

## How It Works

distributeCold() runs at 7AM Cairo:
→ Reads masterLeads (Pending leads)
→ Assigns leads to senders based on capacity and campaign assignment
→ Fires parallel webhook POSTs to all sender inboxes (50 at a time via fetchAll)
→ Immediately kicks off scheduleDailySends on all senders

scheduleDailySends() on each sender:
→ Generates randomized send schedule within campaign time window
→ Creates one-time trigger for first send
→ processSingleSend() fires, sends email, chains next trigger
→ Repeats until daily limit hit or window closes

distributeFollowups() runs at 7:30AM Cairo:
→ Same pattern but reads leads where nextSentDate <= today
→ Followups thread into original email via RFC 2822 headers

## Trigger Schedule (Master Control)

| Time | Function |
|------|----------|
| 6:00 AM | resetSentToday |
| 6:05 AM | runWarmupIncrease |
| 6:30 AM | scanAllSendersForReplies |
| 7:00 AM | distributeCold |
| 7:05 AM | sendDailyDigest |
| 7:30 AM | distributeFollowups |
| Every 20 min | processStagingUpdates |
| 11:00 PM | scanAllSendersForBounces |
| 11:30 PM | runSyncStats |
| 11:55 PM | runDailyReset |

Each sender also has:
- Daily `scheduleDailySends` trigger (fires before window opens)
- Daily `scheduleFollowupSends` trigger (fires before window opens)
- Chained one-time `processSingleSend` triggers during sending window

## Scale

- 100-300 sender accounts
- 10,000-20,000 leads per rotation
- 3-6 simultaneous campaigns
- 2,500-7,500 emails per day at full scale
- fetchAll batched at 50 per call
- Checkpoint/continuation handles 6-minute execution limit

## Key Design Patterns

- **Chained one-time triggers** — never hits Apps Script 20-trigger limit
- **stagingUpdates buffer** — prevents masterLeads write collisions
- **flexibleAllocation** — unused cold slots flow to followups and vice versa
- **Pre-flight window check** — orphan triggers self-delete without touching rows
- **Wasted slot recovery** — rounding losses go to highest priority side
- **23-hour max delay cap** — replaced original 6-hour cap that caused wrong firing times

## API Endpoints

All endpoints accept POST requests with JSON body to Master Control webApp URL.

### Campaigns
| Action | Description |
|--------|-------------|
| createCampaign | Create new campaign with auto-generated ID |
| getCampaigns | List all campaigns with stats and sequence steps |
| updateCampaign | Edit copy, status, window settings |
| deleteCampaign | Remove campaign rows |
| getCampaignStats | Lead status breakdown per campaign |
| sendTestEmail | Send plain text test for any campaign/step |

### Leads
| Action | Description |
|--------|-------------|
| uploadLeads | Bulk upload with duplicate + blocklist check |
| getLeads | Paginated lead list with filters |

### Senders
| Action | Description |
|--------|-------------|
| getSenders | All senders with stats |
| updateSender | Edit sender + sync to Config sheet |
| updateSenderConfig | Bulk update config on all or specific senders |
| updateSendersByTag | Update config for all senders with a tag |
| syncAllSenderConfigs | Push senderAccounts values to all sender Config sheets |
| getSendersByTag | List senders with a specific tag |
| getAllTags | List all unique tags with sender counts |
| assignTag | Assign tag to senders (multi-tag supported, comma separated) |
| removeTag | Remove specific tag from senders |
| getSystemStatus | Live status across all senders |
| getFullSystemReport | Everything in one payload |

### Operations
| Action | Description |
|--------|-------------|
| distributeCold | Run cold distribution now |
| distributeFollowups | Run followup distribution now |
| kickoffAllSenders | Start all cold chains immediately |
| kickoffAllFollowups | Start all followup chains immediately |
| purgeAllInboxes | Clear all inboxes + reset leads correctly |
| systemBoot | Set up all triggers on master + all senders |
| emergencyStop | Pause everything instantly |
| resumeAll | Resume everything |
| resetSentToday | Zero all daily counters |
| cleanupAllTriggers | Kill all orphan triggers system-wide |
| runWarmupIncrease | Manually trigger a ramp up increase |

### Analytics
| Action | Description |
|--------|-------------|
| getDailySummary | Last N days of send stats |
| sendDailyDigest | Trigger digest email manually |

### System
| Action | Description |
|--------|-------------|
| validateSystem | Full health check with pass/fail report |
| healthCheck | Quick ping with active sender count |
| getMasterTriggers | List all master control triggers |
| archiveCampaignLeads | Move campaign leads to new spreadsheet |
| resetOrphanedLeads | Reset stuck leads back to correct state |
| addToBlocklist | Add emails or domains to global blocklist |
| removeFromBlocklist | Remove from blocklist |
| getBlocklist | View full blocklist |

## Sheets Structure

### Master Control Spreadsheet
- `senderAccounts` — all sender accounts with config and stats
- `Campaigns` — campaign rows (one per sequence step)
- `Settings` — global key-value configuration
- `stagingUpdates` — buffer for masterLeads writes
- `dailySummary` — daily send totals
- `Blocklist` — blocked emails and domains

### Leads Database Spreadsheet (separate)
- `masterLeads` — all leads with status tracking

### Each Sender Spreadsheet
- `Config` — sender-specific key-value configuration
- `dailyInbox` — cold email queue
- `followupInbox` — followup email queue

## Known Constraints

- Apps Script 6-minute execution limit
  → Handled via checkpoint/continuation mechanism in distributor.gs
- fetchAll limited to 50 parallel requests per call
  → All fan-outs batched in groups of 50
- CacheService 100KB limit
  → Checkpoints store campaign indices only, not lead data
- stagingUpdates prevents write collisions on masterLeads
- Each sender runs under its own Google account (separate quotas)
- timeapi.io used for timezone-accurate scheduling