/* globals app, NothingEnum */
(function (app) {
    'use strict';

    var document,
        style,
        stylePatterns = {
            'Species Life': {
                '[0-9]+': 'http://lsid.speciesfile.org/$0'
            },
            'WoRMS': {
                '[0-9]+': 'http://www.marinespecies.org/aphia.php?p=taxdetails&id=$0'
            },
            'GenBank': {
                '([A-Z]{1,2}.+\\d+)': 'http://www.ncbi.nlm.nih.gov/nuccore/$1'
            },
            'MycoBank': {
                '(.+)': 'http://www.mycobank.org/MycoTaxo.aspx?Link=T&Rec=$1'
            },
            'FungalNames': {
                '.*?(\\d+)': 'http://fungalinfo.im.ac.cn/fungalname/fungalexample.html?fn=$1'
            },
            'IndexFungorum': {
                'IF(\\d+)': 'http://www.indexfungorum.org/names/NamesRecord.asp?RecordID=$1'
            },
            'Herbimi': {
                '[0-9]+': 'http://www.herbimi.info/herbimi/specimen.htm?imi=$0'
            },
            'Morphbank Collection': {
                '.+': 'http://www.morphbank.net/Browse/ByImage/?tsn=$0'
            },
            'Morphbank': {
                '[0-9]{6,6}': 'http://www.morphbank.net/?id=$0'
            },
            'Ohio Uni LSID': {
                '.+': 'http://lsid.tdwg.org/$0'
            },
            'TreeBase': {
                '.+': 'http://treebase.org/treebase-web/search/study/summary.html?id=$0'
            },
            'Arctos': {
                '.+': 'http://arctos.database.museum/guid/$0'
            },
            'BM barcodes': { // Botanical garden
                '.+': 'http://www.kew.org/collections/wallich/explore/$0.htm'
            },
            'BOLD_Process_ID': {
                '(.+)': 'http://www.boldsystems.org/index.php/Public_RecordView?processid=$1'
            },
            'BOLD_BIN': {
                '(BOLD)?\\W*(.+)': 'http://boldsystems.org/index.php/Public_BarcodeCluster?clusteruri=BOLD:$2'
            },
            'AntWeb': {
                '.+': 'http://www.antweb.org/specimen.do?name=$0'
            },
            'CASENT': {
                '(CASENT|CASTYPE|LACMENT)\\s*(\\d{5,}[A-Z0-9-]*)': 'http://data.antweb.org/specimen/$1$2',
                '([A-Z]+\\s*)?(\\d{5,}[A-Z0-9-]*)': 'http://data.antweb.org/specimen/CASENT$2'
            },
            'OSUC': {
                '(\\d+)': 'http://hol.osu.edu/spmInfo.html?id=OSUC%20$1'
            },
            'USNMENT': {
                '(?:USNMENT\\s*)?([a-z0-9_]+)': 'http://hol.osu.edu/spmInfo.html?id=USNMENT$1'
            },
            'CNC': {
                '(?:CNC)?([0-9]+)': 'http://hol.osu.edu/spmInfo.html?id=CNC$1'
            },
            'UAM': {
                '(UAM:\\w+:\\d+)': 'http://arctos.database.museum/guid/$1'
            },
            'EMEC': {
                '(EMEC\\s*)?(\\d{6,6})': 'http://essigdb.berkeley.edu/cgi/eme_query?table=eme&bnhm_id=EMEC$2&one=T&OK2SHOWPRIVATE='
            },
            'GBIFKey': {
                '(.+)': 'http://www.gbif.org/dataset/$1'
            },
            'UBC A0000 CODES': {
                //'(\\w+)': 'http://bridge.botany.ubc.ca/herbarium/details.php?db=ubcalgae.fmp12&layout=ubcalgae_web_details&recid=210219&ass_num=$1',
            },
            'UK Grid Reference Finder': {
                '.+': 'http://gridreferencefinder.com/?gr=$0'
            },
            'Avibase': {
                '.+': 'http://avibase.bsc-eoc.org/species.jsp?lang=EN&avibaseid=$0&sec=summary&ssver=1'
            }
        },
        pattern,
        harvestPatterns = {
            '(?i)\\bdoi:?\\s*\\d\\S+': {
                '\\d\\S+': 'https://doi.org/$0'
            },
            '(?i)(https?|s?ftp)://\\S+': {
                '.+': '$0'
            },
            '(?i)(?<!://)(www)(\\S+)': {
                '.+': 'http://$0'
            }
        },
        patterns = {
            '(?i)\\b[A-Z0-9\._%+-]+@[A-Z0-9\.-]+\\.[A-Z]{2,4}\\b': {
                '.*': 'mailto:$0'
            },
            '(?<=\\s)urn:lsid:ipni.org:names:\\S+': {
                '.+': 'http://ipni.org/$0'
            },
            '(?<=\\s)urn:lsid:biosci.ohio-state.edu:osuc_pubs:\\d+\\b': {
                '.+': 'http://lsid.tdwg.org/$0'
            },
            '(?<=\\s)urn:lsid:biocol.org:col::\\S+': {
                '.+': 'http://biocol.org/$0'
            },
            '(?i)\\bPMID:?\\s*\\d+': {
                '(?i)PMID:?\\s*(\\d+)': 'http://www.ncbi.nlm.nih.gov/pubmed/$1'
            },
            '(?i)\\bPMCID:?\\s*\\d+': {
                '(?i)PMCID:?\\s*(\\d+)': 'http://www.ncbi.nlm.nih.gov/pmc/articles/PMC$1'
            }
        };

    if (app.documents.length > 0) {
        try {
            document = app.activeDocument;

            removeHyperlinks(document);

            for (pattern in harvestPatterns) {
                try {
                    makeLinksByHarvestPattern(app, document, pattern, harvestPatterns[pattern]);
                } catch (e) {
                    // $.writeln(e);
                }
            }

            for (pattern in patterns) {
                try {
                    makeLinksByPattern(app, document, pattern, patterns[pattern]);
                } catch (e) {
                    // $.writeln(e);
                }
            }

            makeCopyrightsHyperlink(app, document);

            for (var style in stylePatterns) {
                try {
                    makeLinksByCharacterStyle(app, document, style, stylePatterns[style]);
                } catch (e) {
                    // $.writeln(e);
                }
            }

            //alert('DONE! All links have been created!');
        } catch (e) {
            alert(e);
        }
    } else {
        alert('Please open a document');
    }

    function setGrepPreferences(app) {
        app.findGrepPreferences = NothingEnum.nothing;
        app.changeGrepPreferences = NothingEnum.nothing;
        app.findChangeGrepOptions.includeFootnotes = true;
        app.findChangeGrepOptions.includeHiddenLayers = false;
        app.findChangeGrepOptions.includeLockedLayersForFind = true;
        app.findChangeGrepOptions.includeLockedStoriesForFind = true;
        app.findChangeGrepOptions.includeMasterPages = true;
    }

    function setTextPreferences(app) {
        app.findTextPreferences = NothingEnum.nothing;
        app.changeTextPreferences = NothingEnum.nothing;
    }

    function removeHyperlinks(document) {
        var i, item, len, hyperlinks, textSources;

        try {
            textSources = document.hyperlinkTextSources.everyItem().getElements();
            for (i = 0, len = textSources.length; i < len; i += 1) {
                try {
                    item = textSources[i];
                    item.remove();
                } catch (e) {
                    // Do nothing
                }
            }
        } catch (e) {
            throw 'Can not remove text sources';
        }

        try {
            hyperlinks = document.hyperlinks;
            for (i = 0, len = hyperlinks.length; i < len; i += 1) {
                try {
                    item = hyperlinks.item(i);
                    item.remove();
                } catch (e) {
                    // Do nothing
                }
            }
        } catch (e) {
            throw 'Can not remove hyperlinks';
        }
    }

    function makeLinksByPattern(app, document, grepPattern, patterns) {
        var foundItems, numberOffoundItems, pattern, i;

        setGrepPreferences(app);

        app.findGrepPreferences.findWhat = grepPattern;

        foundItems = document.findGrep();

        numberOffoundItems = foundItems.length;
        if (numberOffoundItems < 1) {
            return;
        }

        for (i = 0; i < foundItems.length; i += 1) {
            for (pattern in patterns) {
                processSingleItemFoundByGrep(document, foundItems[i], i, pattern, patterns[pattern]);
            }
        }
    }

    function makeLinksByHarvestPattern(app, document, grepHarvestPattern, patterns) {
        var i, len, k, array = [], item, items = {},
            foundItem, foundItems, numberOffoundItems, pattern;

        setGrepPreferences(app);
        app.findGrepPreferences.findWhat = grepHarvestPattern;

        foundItems = document.findGrep();

        numberOffoundItems = foundItems.length;
        if (numberOffoundItems < 1) {
            return;
        }

        for (i = 0; i < numberOffoundItems; i += 1) {
            foundItem = foundItems[i];
            if (foundItem) {
                item = foundItem.contents.toString().replace(/^[\s<>]+|[\s<>,;:\.\(\)\[\]\{\}]+$/g, '');
                if (item !== '' && !(item in items)) {
                    items[item] = item;
                }
            }
        }

        for (item in items) {
            array.push(item);
        }

        array.sort(function (a, b) {
            return a.length < b.length;
        });

        for (k = 0, len = array.length; k < len; k += 1) {
            item = array[k];
            setTextPreferences(app);
            app.findTextPreferences.findWhat = item;
            foundItems = document.findText();
            numberOffoundItems = foundItems.length;
            if (numberOffoundItems > 0) {
                for (i = 0; i < numberOffoundItems; i += 1) {
                    foundItem = foundItems[i];
                    for (pattern in patterns) {
                        processSingleItemFoundByGrep(document, foundItem, item + '#' + i, pattern, patterns[pattern]);
                    }
                }
            }
        }
    }

    function makeCopyrightsHyperlink(app, document) {
        var foundItems,
            k,
            i,
            len,
            search = {
                'Creative Commons Attribution License \\(CC BY 4\\.0\\)': 'http://creativecommons.org/licenses/by/4.0/'
            };

        setGrepPreferences(app);

        for (k in search) {
            app.findGrepPreferences.findWhat = k;
            foundItems = document.findGrep();

            len = foundItems.length;
            if (len == 0) {
                alert('Found no items to create hyperlinks ' + k);
            }

            for (i = 0; i < len; i += 1) {
                makeHyperlink(document, foundItems[i], search[k], i);
            }
        }
    }

    function makeLinksByCharacterStyle(app, document, characterStyle, patterns) {
        var foundItems, numberOffoundItems, pattern, i;

        if (patterns == null) {
            throw 'Patterns are null';
        }

        if (characterStyle == null || typeof characterStyle !== 'string' || characterStyle.length < 1) {
            throw 'CharacterStyle should be valid string';
        }

        setGrepPreferences(app);

        app.findGrepPreferences.appliedCharacterStyle = characterStyle;

        for (pattern in patterns) {
            app.findGrepPreferences.findWhat = pattern;

            foundItems = document.findGrep();

            numberOffoundItems = foundItems.length;
            if (numberOffoundItems < 1) {
                continue;
            }

            for (i = 0; i < numberOffoundItems; i += 1) {
                processSingleItemFoundByGrep(document, foundItems[i], i, pattern, patterns[pattern]);
            }
        }
    }

    function processSingleItemFoundByGrep(document, foundItem, seed, pattern, urlPlaceholder) {
        var i, len, url,
            re = RegExp(pattern, 'i'),
            matches = re.exec(foundItem.contents);

        if (matches == null) {
            return;
        }

        url = urlPlaceholder;
        for (i = 0, len = matches.length; i < len; i += 1) {
            url = url.replace('$' + i, matches[i]);
        }

        makeHyperlink(document, foundItem, url, seed);
    }

    function makeHyperlink(document, foundItem, url, seed) {
        var hyperlinkDestination, hyperlinkTextSource, hyperlink;
        try {
            hyperlinkDestination = document.hyperlinkURLDestinations.add(url, {
                destinationURL: encodeURI(decodeURI(url))
            });

            hyperlinkTextSource = document.hyperlinkTextSources.add(foundItem);

            hyperlink = document.hyperlinks.add(hyperlinkTextSource, hyperlinkDestination);
            hyperlink.visible = true;
            hyperlink.name = foundItem.contents.replace(/\s+/g, '_') + ' ' + seed;
        } catch (e) {
            //$.writeln(e);
        }
    }

}(app));
