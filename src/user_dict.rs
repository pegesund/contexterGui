//! User dictionary backed by redb.
//! Stores user-added words that should not be flagged as spelling errors.
//! Words get wildcard readings so they pass grammar checks.

use redb::{Database, ReadableTable, TableDefinition};
use mtag::types::{Reading, Pos, Tag};
use std::path::Path;

const WORDS_TABLE: TableDefinition<&str, &str> = TableDefinition::new("user_words");

pub struct UserDict {
    db: Database,
}

impl UserDict {
    pub fn open<P: AsRef<Path>>(path: P) -> Result<Self, Box<dyn std::error::Error>> {
        let db = Database::create(path)?;
        // Ensure table exists
        let txn = db.begin_write()?;
        { let _ = txn.open_table(WORDS_TABLE)?; }
        txn.commit()?;
        Ok(Self { db })
    }

    pub fn add_word(&self, word: &str) -> Result<(), Box<dyn std::error::Error>> {
        let lower = word.to_lowercase();
        let txn = self.db.begin_write()?;
        {
            let mut table = txn.open_table(WORDS_TABLE)?;
            table.insert(lower.as_str(), "")?;
        }
        txn.commit()?;
        Ok(())
    }

    pub fn remove_word(&self, word: &str) -> Result<(), Box<dyn std::error::Error>> {
        let lower = word.to_lowercase();
        let txn = self.db.begin_write()?;
        {
            let mut table = txn.open_table(WORDS_TABLE)?;
            table.remove(lower.as_str())?;
        }
        txn.commit()?;
        Ok(())
    }

    pub fn has_word(&self, word: &str) -> bool {
        let lower = word.to_lowercase();
        let txn = match self.db.begin_read() {
            Ok(t) => t,
            Err(_) => return false,
        };
        let table = match txn.open_table(WORDS_TABLE) {
            Ok(t) => t,
            Err(_) => return false,
        };
        table.get(lower.as_str()).ok().flatten().is_some()
    }

    pub fn list_words(&self) -> Vec<String> {
        let txn = match self.db.begin_read() {
            Ok(t) => t,
            Err(_) => return vec![],
        };
        let table = match txn.open_table(WORDS_TABLE) {
            Ok(t) => t,
            Err(_) => return vec![],
        };
        let mut words = Vec::new();
        let iter = match table.iter() {
            Ok(i) => i,
            Err(_) => return vec![],
        };
        for entry in iter {
            if let Ok((key, _)) = entry {
                words.push(key.value().to_string());
            }
        }
        words
    }

    /// Generate wildcard readings for a user-added word.
    /// Covers noun (all genders), adjective, verb, and adverb
    /// so the word can fill any grammatical slot in Prolog.
    pub fn wildcard_readings(word: &str) -> Vec<Reading> {
        let lemma = word.to_lowercase();
        vec![
            // Noun — masculine singular indefinite
            Reading { lemma: lemma.clone(), pos: Pos::Subst, tags: vec![Tag::Normert, Tag::Appell, Tag::Mask, Tag::Ent, Tag::Ub] },
            // Noun — feminine singular indefinite
            Reading { lemma: lemma.clone(), pos: Pos::Subst, tags: vec![Tag::Normert, Tag::Appell, Tag::Fem, Tag::Ent, Tag::Ub] },
            // Noun — neuter singular indefinite
            Reading { lemma: lemma.clone(), pos: Pos::Subst, tags: vec![Tag::Normert, Tag::Appell, Tag::Noyt, Tag::Ent, Tag::Ub] },
            // Noun — plural indefinite
            Reading { lemma: lemma.clone(), pos: Pos::Subst, tags: vec![Tag::Normert, Tag::Appell, Tag::Fl, Tag::Ub] },
            // Adjective — positive indefinite
            Reading { lemma: lemma.clone(), pos: Pos::Adj, tags: vec![Tag::Normert, Tag::Pos, Tag::Ub] },
            // Adjective — definite
            Reading { lemma: lemma.clone(), pos: Pos::Adj, tags: vec![Tag::Normert, Tag::Pos, Tag::Be] },
            // Verb — infinitive
            Reading { lemma: lemma.clone(), pos: Pos::Verb, tags: vec![Tag::Normert, Tag::Inf] },
            // Verb — present
            Reading { lemma: lemma.clone(), pos: Pos::Verb, tags: vec![Tag::Normert, Tag::Pres] },
            // Verb — past
            Reading { lemma: lemma.clone(), pos: Pos::Verb, tags: vec![Tag::Normert, Tag::Pret] },
            // Adverb
            Reading { lemma: lemma.clone(), pos: Pos::Adv, tags: vec![Tag::Normert] },
        ]
    }
}
