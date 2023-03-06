using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.Linq;
using System.Threading;
using System.Threading.Tasks;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Media;
using ChaosRecipeEnhancer.UI.Filter;
using ChaosRecipeEnhancer.UI.Properties;
using ChaosRecipeEnhancer.UI.View;

namespace ChaosRecipeEnhancer.UI.Model
{
    public static class Data
    {
        public static ActiveItemTypes ActiveItems { get; set; } = new ActiveItemTypes();
        public static ActiveItemTypes PreviousActiveItems { get; set; }
        public static MediaPlayer Player { get; set; } = new MediaPlayer();
        public static MediaPlayer PlayerSet { get; set; } = new MediaPlayer();
        public static int SetAmount { get; set; } = 0;
        public static int SetTargetAmount { get; set; }
        public static List<ItemSet> ItemSetList { get; set; }
        public static List<ItemSet> ItemSetListHighlight { get; set; } = new List<ItemSet>();
        public static ItemSet ItemSetShaper { get; set; }
        public static ItemSet ItemSetElder { get; set; }
        public static ItemSet ItemSetWarlord { get; set; }
        public static ItemSet ItemSetCrusader { get; set; }
        public static ItemSet ItemSetRedeemer { get; set; }
        public static ItemSet ItemSetHunter { get; set; }
        public static CancellationTokenSource cs { get; set; } = new CancellationTokenSource();
        public static CancellationToken CancelationToken { get; set; } = cs.Token;

        public static void GetSetTargetAmount(StashTab stash)
        {
            if (Settings.Default.FullSetThreshold > 0)
            {
                SetTargetAmount = Settings.Default.FullSetThreshold;
            }
            else
            {
                if (stash.Quad)
                    SetTargetAmount += 16;
                else
                    SetTargetAmount += 4;
            }
        }

        private static void GenerateInfluencedItemSets()
        {
            ItemSetShaper = new ItemSet { InfluenceType = "shaper" };
            ItemSetElder = new ItemSet { InfluenceType = "elder" };
            ItemSetWarlord = new ItemSet { InfluenceType = "warlord" };
            ItemSetHunter = new ItemSet { InfluenceType = "hunter" };
            ItemSetCrusader = new ItemSet { InfluenceType = "crusader" };
            ItemSetRedeemer = new ItemSet { InfluenceType = "redeemer" };
        }

        private static void GenerateItemSetList()
        {
            var ret = new List<ItemSet>();
            for (var i = 0; i < SetTargetAmount; i++) ret.Add(new ItemSet());

            ItemSetList = ret;
            Trace.WriteLine(ItemSetList.Count, "[Data:GenerateItemSetList()]: item set list count");
            if (Settings.Default.ExaltedShardRecipeTrackingEnabled) GenerateInfluencedItemSets();
        }

        // TODO: This should maybe be in the Item class?
        private static double GetItemDistanceSq(Item itemA, Item itemB)
        {
            if (itemA.StashTabIndex != itemB.StashTabIndex)
                return ItemSet.StashTabDistanceSq;

            return Math.Pow(itemA.x - itemB.x, 2) + Math.Pow(itemA.y - itemB.y, 2);
        }

        private static Item SelectNearestItem(Item from, string itemType, bool chaosItems)
        {
            Item selectedItem = null;
            var selectedDistSq = double.PositiveInfinity;

            foreach (var stashTab in StashTabList.StashTabs)
            {
                foreach (var item in (List<Item>)Utility.GetPropertyValue(stashTab, chaosItems ? "ItemListChaos" : "ItemList"))
                {
                    if (item.ItemType != itemType)
                        continue;

                    if (from == item)
                        continue;

                    if (from == null)
                        return item;

                    var distSq = GetItemDistanceSq(from, item);
                    if (distSq < selectedDistSq)
                    {
                        selectedItem = item;
                        selectedDistSq = distSq;
                    }
                }
            }

            return selectedItem;
        }
        
        private static void AddItem(ItemSet set, Item item, bool chaosItems)
        {
            set.AddItem(item);
            var tab = GetStashTabFromItem(item);
            ((List<Item>)Utility.GetPropertyValue(tab, chaosItems ? "ItemListChaos" : "ItemList")).Remove(item);
            tab.RemoveItemFromGrid(item);
        }

        private static Item AddWeaponsToSet(ItemSet set, bool chaosItems, Item previousItem)
        {
            // Gather weapon set options: Either one Two-Hander or two One-Handers

            Item twoHander = SelectNearestItem(previousItem, "TwoHandWeapons", chaosItems);

            Item oneHanderA = SelectNearestItem(previousItem, "OneHandWeapons", chaosItems);
            Item oneHanderB = null;
            if (oneHanderA != null)
                oneHanderB = SelectNearestItem(oneHanderA, "OneHandWeapons", chaosItems);

            // No options?
            if (twoHander == null && (oneHanderA == null || oneHanderB == null))
                return null;

            // Does only one option exist?
            if (twoHander == null)
            {
                AddItem(set, oneHanderA, chaosItems);
                AddItem(set, oneHanderB, chaosItems);
                return oneHanderB;
            }
            else if (oneHanderA == null || oneHanderB == null)
            {
                AddItem(set, twoHander, chaosItems);
                return twoHander;
            }

            // If this is the first pick then we can pick arbitrarily
            if (previousItem == null)
            {
                AddItem(set, twoHander, chaosItems);
                return twoHander;
            }

            // Otherwise, consider distance when choosing:
            var twoHanderDistSq = Math.Sqrt(GetItemDistanceSq(previousItem, twoHander));
            var oneHandersDistSq = Math.Sqrt(GetItemDistanceSq(previousItem, oneHanderA)) + Math.Sqrt(GetItemDistanceSq(oneHanderA, oneHanderB));
            if (twoHanderDistSq <= oneHandersDistSq)
            {
                AddItem(set, twoHander, chaosItems);
                return twoHander;
            }
            else
            {
                AddItem(set, oneHanderA, chaosItems);
                AddItem(set, oneHanderB, chaosItems);
                return oneHanderB;
            }
        }

        private static Item AddItemTypeToSet(ItemSet set, string itemType, bool chaosItems, Item previousItem)
        {
            if (itemType == "Weapons")
                return AddWeaponsToSet(set, chaosItems, previousItem);

            var nearest = SelectNearestItem(previousItem, itemType, chaosItems);
            if (nearest != null)
                AddItem(set, nearest, chaosItems);

            return nearest;
        }

        // Greedy O(n), for n items, algorithm to find a "decent" take-out order, but not the optimal
        private static bool FillSingleItemSet(ItemSet set, bool chaosItems, List<Tuple<string, int>> counts)
        {
            // We use a greedy algorithm:
            // 1. Take items in order of least available (dictated by counts).
            // 2. Take the next item such that it minimizes the distance to the previous.
            // 3. Sort the resulting set to furhter minimize takeout distance (i.e. remove inefficiencies due to ordering in 1.)

            Item previousItem = null;
            foreach (var count in counts)
            {
                if (count.Item1 == "Rings")
                {
                    var item = AddItemTypeToSet(set, count.Item1, chaosItems, previousItem);
                    if (item == null)
                        return false;
                    item = AddItemTypeToSet(set, count.Item1, chaosItems, item);
                    if (item == null)
                        return false;
                    previousItem = item;
                }
                else
                {
                    var item = AddItemTypeToSet(set, count.Item1, chaosItems, previousItem);
                    if (item == null)
                        return false;
                    previousItem = item;
                }
            }

            SortTakeoutOrder(set);

            return true;
        }

        private static void SortTakeoutOrder(ItemSet set)
        {
            for (int i = 0; i < set.ItemList.Count - 1; ++i)
            {
                var from = set.ItemList[i];
                var closestItem = set.ItemList.Skip(i + 1).OrderBy(it => GetItemDistanceSq(from, it)).First();
                var closestIndex = set.ItemList.IndexOf(closestItem);

                set.ItemList[closestIndex] = set.ItemList[i + 1];
                set.ItemList[i + 1] = closestItem;                
            }
        }

        // Expanding square flooding algorithm for finding items to take out that minimize distance to a given
        // starting item (which is determiend by selecting item type with lowest count)
        private static bool FillSingleItemSetAlt(ItemSet set, bool chaosItems, List<Tuple<string, int>> counts)
        {
            var firstItem = AddItemTypeToSet(set, counts[0].Item1, chaosItems, null);
            if (firstItem == null)
                return false;
            var floodTab = GetStashTabFromItem(firstItem);
            var centerX = (int)Math.Floor(firstItem.x + 0.5 * firstItem.w);
            var centerY = (int)Math.Floor(firstItem.y + 0.5 * firstItem.h);

            var needed = new List<string>() { "Rings", "Rings", "Amulets", "Belts", "BodyArmours", "Gloves", "Helmets", "Boots", "Weapons" };
            needed.Remove(firstItem.ItemType);

            Item previousOneHander = null;
            Item previousItem = null;

            int topX = centerX - 1, bottomX = centerX + 1, topY = centerY - 1, bottomY = centerY + 1;
            var sz = floodTab.Quad ? 24 : 12;
            while (topX >= 0 || topY >= 0 || bottomX <= sz || bottomY <= sz)
            {
                // We loop the inner border of our expanding square, note that some of it may be outside the stash tab
                for (int y = topY; y <= bottomY; ++y)
                {
                    for (int x = topX; x <= bottomX; /*empty*/)
                    {
                        Item item = null;
                        if (x >= 0 && y >= 0 && x < sz && y < sz)
                            item = floodTab.ItemGrid[x, y];

                        if (item != null)
                        {
                            if (item.ItemType == "TwoHandWeapons" && needed.Contains("Weapons"))
                            {
                                AddItem(set, item, chaosItems);
                                needed.Remove("Weapons");
                                previousItem = item;
                            }
                            else if (item.ItemType == "OneHandWeapons" && needed.Contains("Weapons"))
                            {
                                if (previousOneHander == null)
                                {
                                    previousOneHander = item;
                                }
                                else if (previousOneHander != item)
                                {
                                    AddItem(set, previousOneHander, chaosItems);
                                    AddItem(set, item, chaosItems);
                                    needed.Remove("Weapons");
                                    previousItem = item;
                                }
                            }
                            else if (needed.Contains(item.ItemType))
                            {
                                AddItem(set, item, chaosItems);
                                needed.Remove(item.ItemType);
                                previousItem = item;
                            }
                        }

                        if (y == topY || y == bottomY)
                            ++x;
                        else
                            x += (bottomX - topX);
                    }
                }

                // TODO: Early exist if needed is empty

                --topX; --topY; ++bottomX; ++bottomY;
            }

            foreach (var need in needed)
            {
                previousItem = AddItemTypeToSet(set, need, chaosItems, previousItem);
                if (previousItem == null)
                    return false;
            }

            // Take top-leftmost item out first (TODO: We should just show overlay for all items to take out at once)
            var topLeftItem = set.ItemList.Where(it => it.y == set.ItemList.Min(ite => ite.y)).OrderBy(it => it.x).First();
            var topLeftIndex = set.ItemList.IndexOf(topLeftItem);
            set.ItemList[topLeftIndex] = set.ItemList[0];
            set.ItemList[0] = topLeftItem;
            SortTakeoutOrder(set);

            return true;
        }

        private static int FillItemSets(List<Tuple<string, int>> chaosItemCounts, List<Tuple<string, int>> regalItemCounts)
        {
            // TODO: What is the meaning of Settings.Default.DoNotPreserveLowItemLevelGear
            // TODO: Remove all traces of honoring given order, minimizing distance is preferable

            chaosItemCounts.Sort((x, y) => x.Item2.CompareTo(y.Item2));
            regalItemCounts.Sort((x, y) => x.Item2.CompareTo(y.Item2));

            int i = 0;
            if (Settings.Default.ChaosRecipeTrackingEnabled)
            {
                while (i < ItemSetList.Count && FillSingleItemSetAlt(ItemSetList[i], true, chaosItemCounts))
                    ++i;
            }

            if (Settings.Default.RegalRecipeTrackingEnabled)
            {
                while (i < ItemSetList.Count && FillSingleItemSetAlt(ItemSetList[i], false, regalItemCounts))
                    ++i;
            }

            return i;
        }

        private static void FillItemSetsInfluenced()
        {
            foreach (var tab in StashTabList.StashTabs)
            {
                foreach (var i in tab.ItemListShaper)
                {
                    if (ItemSetShaper.EmptyItemSlots.Count == 0) break;

                    ItemSetShaper.AddItem(i);
                }

                foreach (var i in tab.ItemListElder)
                {
                    if (ItemSetElder.EmptyItemSlots.Count == 0) break;

                    ItemSetElder.AddItem(i);
                }

                foreach (var i in tab.ItemListCrusader)
                {
                    if (ItemSetCrusader.EmptyItemSlots.Count == 0) break;

                    ItemSetCrusader.AddItem(i);
                }

                foreach (var i in tab.ItemListWarlord)
                {
                    if (ItemSetWarlord.EmptyItemSlots.Count == 0) break;

                    ItemSetWarlord.AddItem(i);
                }

                foreach (var i in tab.ItemListRedeemer)
                {
                    if (ItemSetRedeemer.EmptyItemSlots.Count == 0) break;

                    ItemSetRedeemer.AddItem(i);
                }

                foreach (var i in tab.ItemListHunter)
                {
                    if (ItemSetHunter.EmptyItemSlots.Count == 0) break;

                    ItemSetHunter.AddItem(i);
                }
            }
        }

        public static async Task CheckActives(SetTrackerOverlayView setTrackerOverlay)
        {
            try
            {
                if (ApiAdapter.FetchError)
                {
                    setTrackerOverlay.WarningMessage = "Fetching Error...";
                    setTrackerOverlay.ShadowOpacity = 1;
                    setTrackerOverlay.WarningMessageVisibility = Visibility.Visible;
                    return;
                }

                if (StashTabList.StashTabs.Count == 0)
                {
                    setTrackerOverlay.WarningMessage = "No Stashtabs found...";
                    setTrackerOverlay.ShadowOpacity = 1;
                    setTrackerOverlay.WarningMessageVisibility = Visibility.Visible;
                    return;
                }

                if (Settings.Default.SoundEnabled) PreviousActiveItems = new ActiveItemTypes(ActiveItems);

                // calculate target amount if user has 0 set in it
                // (e.g. 2 quad tabs queried w 0 set threshold = 24 set threshold)
                // else just stick to the default amount (their defined in settings)
                SetTargetAmount = 0;
                if (StashTabList.StashTabs.Count > 0)
                    foreach (var s in StashTabList.StashTabs)
                        GetSetTargetAmount(s);

                Trace.WriteLine("Calculating Items");
                CalculateItemAmounts(Settings.Default.SetTrackerOverlayItemCounterDisplayMode != 0 ? setTrackerOverlay : null,
                    out var chaosItemCounts, out var regalItemCounts, out var missingItemClasses);

                // generate {SetThreshold} empty sets to be filled
                GenerateItemSetList();

                // proceed to fill those newly created empty sets
                var fullSets = FillItemSets(chaosItemCounts, regalItemCounts);

                // check for full sets/ missing items
                var missingGearPieceForChaosRecipe = false; // TODO

                var filterManager = new CFilterGenerationManager();

                // i need to pass in the missingGearPieceForChaosRecipe
                ActiveItems =
                    await filterManager.GenerateSectionsAndUpdateFilterAsync(missingItemClasses);

                //Trace.WriteLine(fullSets, "full sets");
                setTrackerOverlay.Dispatcher.Invoke(() => { setTrackerOverlay.FullSetsText = fullSets.ToString(); });

                // invoke chaos missing
                if (missingGearPieceForChaosRecipe && !Settings.Default.RegalRecipeTrackingEnabled)
                {
                    setTrackerOverlay.WarningMessage = "Need lower level items!";
                    setTrackerOverlay.ShadowOpacity = 1;
                    setTrackerOverlay.WarningMessageVisibility = Visibility.Visible;
                }

                // invoke exalted recipe ready
                if (Settings.Default.ExaltedShardRecipeTrackingEnabled)
                    if (ItemSetShaper.EmptyItemSlots.Count == 0
                        || ItemSetElder.EmptyItemSlots.Count == 0
                        || ItemSetCrusader.EmptyItemSlots.Count == 0
                        || ItemSetWarlord.EmptyItemSlots.Count == 0
                        || ItemSetHunter.EmptyItemSlots.Count == 0
                        || ItemSetRedeemer.EmptyItemSlots.Count == 0)
                    {
                        setTrackerOverlay.WarningMessage = "Exalted Recipe ready!";
                        setTrackerOverlay.ShadowOpacity = 1;
                        setTrackerOverlay.WarningMessageVisibility = Visibility.Visible;
                    }

                // invoke set full
                if (fullSets == SetTargetAmount && !missingGearPieceForChaosRecipe)
                {
                    setTrackerOverlay.WarningMessage = "Sets full!";
                    setTrackerOverlay.ShadowOpacity = 1;
                    setTrackerOverlay.WarningMessageVisibility = Visibility.Visible;
                }

                Trace.WriteLine(fullSets, "full sets");

                // If the state of any gear slot changed, we play a sound
                if (Settings.Default.SoundEnabled)
                    if (!(PreviousActiveItems.GlovesActive == ActiveItems.GlovesActive
                          && PreviousActiveItems.BootsActive == ActiveItems.BootsActive
                          && PreviousActiveItems.HelmetActive == ActiveItems.HelmetActive
                          && PreviousActiveItems.ChestActive == ActiveItems.ChestActive
                          && PreviousActiveItems.WeaponActive == ActiveItems.WeaponActive
                          && PreviousActiveItems.RingActive == ActiveItems.RingActive
                          && PreviousActiveItems.AmuletActive == ActiveItems.AmuletActive
                          && PreviousActiveItems.BeltActive == ActiveItems.BeltActive))
                        Player.Dispatcher.Invoke(() =>
                        {
                            Trace.WriteLine("Gear Slot State Changed; Playing sound!");
                            PlayNotificationSound();
                        });
            }
            catch (OperationCanceledException ex) when (ex.CancellationToken == CancelationToken)
            {
                Trace.WriteLine("abort");
            }
        }

        public static void CalculateItemAmounts(SetTrackerOverlayView setTrackerOverlay,
            out List<Tuple<string, int>> chaosItems, out List<Tuple<string, int>> regalItems, out HashSet<string> missing)
        {
            chaosItems = new List<Tuple<string, int>>();
            regalItems = new List<Tuple<string, int>>();
            missing = new HashSet<string>();

            // 0: rings
            // 1: amulets
            // 2: belts
            // 3: chests
            // 4: gloves
            // 5: helmets
            // 6: boots
            // 7: two hand weapons
            // 8: one hand weapons
            var regalAmounts = new int[9];
            var chaosAmounts = new int[9];

            if (StashTabList.StashTabs != null)
            {
                Trace.WriteLine("calculating items amount");

                foreach (var tab in StashTabList.StashTabs)
                {
                    Trace.WriteLine("tab amount " + tab.ItemList.Count);
                    Trace.WriteLine("tab amount " + tab.ItemListChaos.Count);

                    if (tab.ItemList.Count > 0)
                        foreach (var i in tab.ItemList)
                        {
                            Trace.WriteLine(i.ItemType);
                            if (i.ItemType == "Rings")
                                regalAmounts[0]++;
                            else if (i.ItemType == "Amulets")
                                regalAmounts[1]++;
                            else if (i.ItemType == "Belts")
                                regalAmounts[2]++;
                            else if (i.ItemType == "BodyArmours")
                                regalAmounts[3]++;
                            else if (i.ItemType == "Gloves")
                                regalAmounts[4]++;
                            else if (i.ItemType == "Helmets")
                                regalAmounts[5]++;
                            else if (i.ItemType == "Boots")
                                regalAmounts[6]++;
                            else if (i.ItemType == "TwoHandWeapons")
                                regalAmounts[7]++;
                            else if (i.ItemType == "OneHandWeapons")
                                regalAmounts[8]++;
                        }

                    if (tab.ItemListChaos.Count > 0)
                        foreach (var i in tab.ItemListChaos)
                        {
                            Trace.WriteLine(i.ItemType);
                            if (i.ItemType == "Rings")
                                chaosAmounts[0]++;
                            else if (i.ItemType == "Amulets")
                                chaosAmounts[1]++;
                            else if (i.ItemType == "Belts")
                                chaosAmounts[2]++;
                            else if (i.ItemType == "BodyArmours")
                                chaosAmounts[3]++;
                            else if (i.ItemType == "Gloves")
                                chaosAmounts[4]++;
                            else if (i.ItemType == "Helmets")
                                chaosAmounts[5]++;
                            else if (i.ItemType == "Boots")
                                chaosAmounts[6]++;
                            else if (i.ItemType == "TwoHandWeapons")
                                chaosAmounts[7]++;
                            else if (i.ItemType == "OneHandWeapons")
                                chaosAmounts[8]++;
                        }
                }

                // Update missing
                if (chaosAmounts[0] / 2 + regalAmounts[0] / 2 < SetTargetAmount) missing.Add("Rings");
                if (chaosAmounts[1] + regalAmounts[1] < SetTargetAmount) missing.Add("Amulets");
                if (chaosAmounts[2] + regalAmounts[2] < SetTargetAmount) missing.Add("Belts");
                if (chaosAmounts[3] + regalAmounts[3] < SetTargetAmount) missing.Add("BodyArmours");
                if (chaosAmounts[4] + regalAmounts[4] < SetTargetAmount) missing.Add("Gloves");
                if (chaosAmounts[5] + regalAmounts[5] < SetTargetAmount) missing.Add("Helmets");
                if (chaosAmounts[6] + regalAmounts[6] < SetTargetAmount) missing.Add("Boots");
                if (chaosAmounts[7] + regalAmounts[7] + chaosAmounts[8] / 2 + regalAmounts[8] / 2 < SetTargetAmount)
                {
                    missing.Add("OneHandWeapons");
                    missing.Add("TwoHandWeapons");
                }

                // Update overlay counter
                if (Settings.Default.SetTrackerOverlayItemCounterDisplayMode == 1 && setTrackerOverlay != null)
                {
                    setTrackerOverlay.RingsAmount = regalAmounts[0] + chaosAmounts[0];
                    setTrackerOverlay.AmuletsAmount = regalAmounts[1] + chaosAmounts[1];
                    setTrackerOverlay.BeltsAmount = regalAmounts[2] + chaosAmounts[2];
                    setTrackerOverlay.ChestsAmount = regalAmounts[3] + chaosAmounts[3];
                    setTrackerOverlay.GlovesAmount = regalAmounts[4] + chaosAmounts[4];
                    setTrackerOverlay.HelmetsAmount = regalAmounts[5] + chaosAmounts[5];
                    setTrackerOverlay.BootsAmount = regalAmounts[6] + chaosAmounts[6];
                    setTrackerOverlay.WeaponsAmount = regalAmounts[7] + regalAmounts[8] + chaosAmounts[7] + chaosAmounts[8];
                }
                else if (Settings.Default.SetTrackerOverlayItemCounterDisplayMode == 2 && setTrackerOverlay != null)
                {
                    setTrackerOverlay.RingsAmount = SetTargetAmount * 2 - regalAmounts[0] - chaosAmounts[0];
                    setTrackerOverlay.AmuletsAmount = SetTargetAmount - regalAmounts[1] - chaosAmounts[1];
                    setTrackerOverlay.BeltsAmount = SetTargetAmount - regalAmounts[2] - chaosAmounts[2];
                    setTrackerOverlay.ChestsAmount = SetTargetAmount - regalAmounts[3] - chaosAmounts[3];
                    setTrackerOverlay.GlovesAmount = SetTargetAmount - regalAmounts[4] - chaosAmounts[4];
                    setTrackerOverlay.HelmetsAmount = SetTargetAmount - regalAmounts[5] - chaosAmounts[5];
                    setTrackerOverlay.BootsAmount = SetTargetAmount - regalAmounts[6] - chaosAmounts[6];
                    setTrackerOverlay.WeaponsAmount = SetTargetAmount * 2 -
                        (chaosAmounts[7] + regalAmounts[7] + (chaosAmounts[8] + regalAmounts[8]) * 2);
                }
            }

            int idx = 0;
            foreach (var itemType in new string[] { "Rings", "Amulets", "Belts", "BodyArmours", "Gloves", "Helmets", "Boots" })
            {
                chaosItems.Add(new Tuple<string, int>(itemType, chaosAmounts[idx]));
                regalItems.Add(new Tuple<string, int>(itemType, regalAmounts[idx]));
                ++idx;
            }

            chaosItems.Add(new Tuple<string, int>("Weapons", chaosAmounts[7] + chaosAmounts[8]));
            regalItems.Add(new Tuple<string, int>("Weapons", regalAmounts[7] + regalAmounts[8]));
        }

        public static void PlayNotificationSound()
        {
            var volume = Settings.Default.Volume / 100.0;
            Player.Volume = volume;
            Player.Position = TimeSpan.Zero;
            Player.Play();
        }

        public static void PlayNotificationSoundSetPicked()
        {
            var volume = Settings.Default.Volume / 100.0;
            PlayerSet.Volume = volume;
            PlayerSet.Position = TimeSpan.Zero;
            PlayerSet.Play();
        }

        public static StashTab GetStashTabFromItem(Item item)
        {
            foreach (var s in StashTabList.StashTabs)
                if (item.StashTabIndex == s.TabIndex)
                    return s;

            return null;
        }

        public static void ActivateNextCell(bool active, InteractiveCell cell, TabControl tabControl)
        {
            if (!active) return;

            var currentlySelectedStashOverlayTabName = tabControl != null
                ? ((TextBlock)((HeaderedContentControl)tabControl.SelectedItem).Header).Text
                : "";

            // activate cell by cell / item by item
            if (Settings.Default.StashTabOverlayHighlightMode == 0)
            {
                foreach (var s in StashTabList.StashTabs.ToList())
                {
                    s.DeactivateItemCells();
                    s.TabHeaderColor = Brushes.Transparent;
                }

                // remove and sound if itemlist empty
                if (ItemSetListHighlight.Count > 0)
                {
                    if (ItemSetListHighlight[0].ItemList.Count == 0)
                    {
                        ItemSetListHighlight.RemoveAt(0);
                        PlayerSet.Dispatcher.Invoke(() => { PlayNotificationSoundSetPicked(); });
                    }
                }
                else
                {
                    if (ItemSetListHighlight.Count > 0)
                        PlayerSet.Dispatcher.Invoke(() => { PlayNotificationSoundSetPicked(); });
                }

                // next item if itemlist not empty
                if (ItemSetListHighlight.Count > 0)
                {
                    if (ItemSetListHighlight[0].ItemList.Count > 0 &&
                        ItemSetListHighlight[0].EmptyItemSlots.Count == 0)
                    {
                        var highlightItem = ItemSetListHighlight[0].ItemList[0];
                        var currentTab = GetStashTabFromItem(highlightItem);

                        if (currentTab != null)
                        {
                            currentTab.ActivateItemCells(highlightItem);

                            // if (tabControl != null)
                            // {
                            //     Trace.WriteLine($"[Data: ActivateNextCell()]: TabControl Current Tab Item {tabControl.SelectedItem}");
                            //     Trace.WriteLine($"[Data: ActivateNextCell()]: TabControl Current Tab Item Header Text {((TextBlock)((HeaderedContentControl)tabControl.SelectedItem).Header).Text}");
                            // }

                            if (currentTab.TabName != currentlySelectedStashOverlayTabName &&
                                Settings.Default.StashTabOverlayHighlightColor != "")
                                currentTab.TabHeaderColor = new SolidColorBrush(
                                    (Color)ColorConverter.ConvertFromString(Settings.Default
                                        .StashTabOverlayHighlightColor));
                            else
                                currentTab.TabHeaderColor = Brushes.Transparent;

                            ItemSetListHighlight[0].ItemList.RemoveAt(0);
                        }
                    }
                }
            }
            // activate set by set
            else if (Settings.Default.StashTabOverlayHighlightMode == 1)
            {
                if (ItemSetListHighlight.Count > 0)
                {
                    Trace.WriteLine(ItemSetListHighlight[0].ItemList.Count, "[Data: ActivateNextCell()]: item list count");
                    Trace.WriteLine(ItemSetListHighlight.Count, "[Data: ActivateNextCell()]: item set list count");

                    // check for full sets
                    if (ItemSetListHighlight[0].EmptyItemSlots.Count == 0)
                    {
                        if (cell != null)
                        {
                            var highlightItem = cell.PathOfExileItemData;
                            var currentTab = GetStashTabFromItem(highlightItem);

                            if (currentTab != null)
                            {
                                currentTab.TabHeaderColor = Brushes.Transparent;
                                currentTab.DeactivateSingleItemCells(cell.PathOfExileItemData);
                                ItemSetListHighlight[0].ItemList.Remove(highlightItem);
                            }
                        }

                        foreach (var i in ItemSetListHighlight[0].ItemList.ToList())
                        {
                            var currTab = GetStashTabFromItem(i);
                            currTab.ActivateItemCells(i);
                        }

                        // mark item order
                        if (ItemSetListHighlight[0] != null)
                        {
                            if (ItemSetListHighlight[0].ItemList.Count > 0)
                            {
                                var currentStashTab = GetStashTabFromItem(ItemSetListHighlight[0].ItemList[0]);
                                currentStashTab.MarkNextItem(ItemSetListHighlight[0].ItemList[0]);
                                currentStashTab.TabHeaderColor = new SolidColorBrush((Color)ColorConverter.ConvertFromString(Settings.Default.StashTabOverlayHighlightColor));

                                // if (tabControl != null)
                                // {
                                //     Trace.WriteLine($"[Data: ActivateNextCell()]: TabControl Current Tab Item {tabControl.SelectedItem}");
                                //     Trace.WriteLine($"[Data: ActivateNextCell()]: TabControl Current Tab Item Header Text {((TextBlock)((HeaderedContentControl)tabControl.SelectedItem).Header).Text}");
                                // }

                            }
                        }

                        // Set has been completed
                        if (ItemSetListHighlight[0].ItemList.Count == 0)
                        {
                            ItemSetListHighlight.RemoveAt(0);

                            // activate next set
                            ActivateNextCell(true, null, null);
                            PlayerSet.Dispatcher.Invoke(() => { PlayNotificationSoundSetPicked(); });
                        }
                    }
                }
            }
            // activate whole set
            else if (Settings.Default.StashTabOverlayHighlightMode == 2)
            {
                //activate all cells at once
                if (ItemSetListHighlight.Count <= 0) return;

                // Why do I switch all of the foreach loops to reference {Some List}.ToList()?
                // REF: https://stackoverflow.com/a/604843
                foreach (var set in ItemSetListHighlight.ToList())
                {
                    if (set.EmptyItemSlots.Count != 0) continue;

                    if (cell == null) continue;

                    var highlightItem = cell.PathOfExileItemData;
                    var currentTab = GetStashTabFromItem(highlightItem);

                    if (currentTab == null) continue;

                    currentTab.DeactivateSingleItemCells(cell.PathOfExileItemData);
                    ItemSetListHighlight[0].ItemList.Remove(highlightItem);
                    
                    var itemsRemainingInStashTab = false;

                    foreach (var item in ItemSetListHighlight.ToList()[0].ItemList.ToList())
                    {
                        if (item.StashTabIndex == currentTab.TabIndex) itemsRemainingInStashTab = true;
                    }

                    if (!itemsRemainingInStashTab) currentTab.TabHeaderColor = Brushes.Transparent;

                    // Set has been completed
                    if (ItemSetListHighlight[0].ItemList.Count != 0) continue;

                    ItemSetListHighlight.RemoveAt(0);

                    // activate next set, if one exists
                    if (ItemSetListHighlight.Count != 0) ActivateNextCell(true, null, null);
                    else PlayerSet.Dispatcher.Invoke(() => { PlayNotificationSoundSetPicked(); });
                }
            }
        }

        public static void PrepareSelling()
        {
            //ClearAllItemOrderLists();
            ItemSetListHighlight.Clear();
            if (ApiAdapter.IsFetching) return;

            if (ItemSetList == null) return;

            foreach (var s in StashTabList.StashTabs)
                s.PrepareOverlayList();

            if (Settings.Default.ExaltedShardRecipeTrackingEnabled)
            {
                if (ItemSetShaper.EmptyItemSlots.Count == 0)
                    ItemSetListHighlight.Add(new ItemSet
                    {
                        ItemList = new List<Item>(ItemSetShaper.ItemList),
                        EmptyItemSlots = new List<string>(ItemSetShaper.EmptyItemSlots)
                    });

                if (ItemSetElder.EmptyItemSlots.Count == 0)
                    ItemSetListHighlight.Add(new ItemSet
                    {
                        ItemList = new List<Item>(ItemSetElder.ItemList),
                        EmptyItemSlots = new List<string>(ItemSetElder.EmptyItemSlots)
                    });

                if (ItemSetCrusader.EmptyItemSlots.Count == 0)
                    ItemSetListHighlight.Add(new ItemSet
                    {
                        ItemList = new List<Item>(ItemSetCrusader.ItemList),
                        EmptyItemSlots = new List<string>(ItemSetCrusader.EmptyItemSlots)
                    });

                if (ItemSetHunter.EmptyItemSlots.Count == 0)
                    ItemSetListHighlight.Add(new ItemSet
                    {
                        ItemList = new List<Item>(ItemSetHunter.ItemList),
                        EmptyItemSlots = new List<string>(ItemSetHunter.EmptyItemSlots)
                    });

                if (ItemSetWarlord.EmptyItemSlots.Count == 0)
                    ItemSetListHighlight.Add(new ItemSet
                    {
                        ItemList = new List<Item>(ItemSetWarlord.ItemList),
                        EmptyItemSlots = new List<string>(ItemSetWarlord.EmptyItemSlots)
                    });

                if (ItemSetRedeemer.EmptyItemSlots.Count == 0)
                    ItemSetListHighlight.Add(new ItemSet
                    {
                        ItemList = new List<Item>(ItemSetRedeemer.ItemList),
                        EmptyItemSlots = new List<string>(ItemSetRedeemer.EmptyItemSlots)
                    });
            }

            foreach (var set in ItemSetList)
            {
                if (set.SetCanProduceChaos || Settings.Default.RegalRecipeTrackingEnabled)
                    ItemSetListHighlight.Add(new ItemSet
                    {
                        ItemList = new List<Item>(set.ItemList),
                        EmptyItemSlots = new List<string>(set.EmptyItemSlots)
                    });
            }
        }
    }
}
