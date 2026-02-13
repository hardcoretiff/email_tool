"""
PST Email Searcher - Legal eDiscovery Tool
============================================
Searches through Outlook .pst email archives to find relevant correspondence
based on a list of companies/contacts and topic-specific keywords.

Built for federal criminal defense case review - specifically designed to
identify emails discussing doctor authorizations, compliance issues, and
termination of business relationships with marketing companies.

Usage:
    python pst_email_searcher.py --pst <path_to_pst> --companies <path_to_companies_list>

Requirements:
    pip install pypff pandas openpyxl python-dateutil
    
    For AI-powered relevance scoring (optional):
    pip install anthropic
"""

import os
import sys
import csv
import json
import argparse
import re
from datetime import datetime
from pathlib import Path
from collections import defaultdict

try:
    import pypff
    HAS_PYPFF = True
except ImportError:
    HAS_PYPFF = False

try:
    import pandas as pd
    HAS_PANDAS = True
except ImportError:
    HAS_PANDAS = False

try:
    from anthropic import Anthropic
    HAS_ANTHROPIC = True
except ImportError:
    HAS_ANTHROPIC = False


# =============================================================================
# KEYWORD CONFIGURATION
# =============================================================================
# These are the terms the tool searches for within emails. They're organized
# by category so the relevance scoring can weight them appropriately.
# You can add/remove/modify these to fit different cases.

KEYWORD_CATEGORIES = {
    "doctor_authorization": {
        "weight": 3,  # highest priority - this is the core issue
        "terms": [
            "doctor authorization", "doctor's authorization", "physician authorization",
            "physician order", "physician's order", "doctor order", "doctor's order",
            "medical authorization", "prior authorization", "prior auth",
            "face-to-face", "face to face", "f2f", "certificate of medical necessity",
            "cmn", "medical necessity", "prescription", "prescribing physician",
            "referring physician", "ordering physician", "written order",
            "signed order", "valid order", "order form", "dme order",
            "documentation", "medical records", "patient records",
            "authorization form", "auth form", "proper authorization",
            "without authorization", "no authorization", "missing authorization",
            "lack of authorization", "unauthorized"
        ]
    },
    "compliance_issues": {
        "weight": 2,  # important context
        "terms": [
            "compliance", "non-compliance", "noncompliance", "compliant",
            "non-compliant", "violation", "violations", "issue", "issues",
            "problem", "problems", "concern", "concerns", "complaint",
            "audit", "investigation", "inquiry", "review", "oversight",
            "regulation", "regulatory", "medicare", "medicaid", "cms",
            "oig", "fraud", "abuse", "kickback", "anti-kickback",
            "false claim", "false claims", "billing issue", "billing problem",
            "improper", "irregular", "questionable", "suspicious",
            "cutting corners", "shortcut", "shortcuts", "not following",
            "failed to", "failure to", "neglected", "negligent"
        ]
    },
    "termination_language": {
        "weight": 3,  # highest priority - proves client took action
        "terms": [
            "terminate", "terminated", "termination", "terminating",
            "fire", "fired", "firing", "let go", "letting go",
            "end the relationship", "ending the relationship",
            "discontinue", "discontinued", "discontinuing",
            "cancel", "cancelled", "canceled", "cancellation",
            "sever", "severed", "severing", "part ways", "parting ways",
            "no longer working", "no longer work with",
            "cease", "ceased", "stop working with", "stopped working",
            "contract termination", "end of contract", "breach of contract",
            "effective immediately", "last day", "final notice",
            "transition", "transitioning", "replacement", "replacing",
            "new company", "new vendor", "new marketing", "different company",
            "moving on", "move on", "moved on"
        ]
    },
    "marketing_relationship": {
        "weight": 1,  # supporting context
        "terms": [
            "marketing", "marketing company", "marketing firm",
            "marketing agreement", "marketing contract", "referral",
            "referrals", "lead", "leads", "patient leads", "patient referral",
            "sales", "sales rep", "representative", "account manager",
            "partnership", "partner", "vendor", "contractor", "subcontractor",
            "agreement", "contract", "scope of work", "sow",
            "commission", "compensation", "payment", "invoice"
        ]
    }
}


class EmailRecord:
    """Represents a single parsed email with all its metadata and content."""
    
    def __init__(self):
        self.subject = ""
        self.sender = ""
        self.sender_email = ""
        self.recipients = ""
        self.cc = ""
        self.bcc = ""
        self.date = None
        self.date_str = ""
        self.body = ""
        self.folder_path = ""
        self.has_attachments = False
        self.attachment_names = []
        self.message_id = ""
        
        # analysis fields (populated during search)
        self.relevance_score = 0
        self.matched_keywords = []
        self.matched_categories = []
        self.matched_companies = []
        self.ai_relevance_score = None
        self.ai_summary = ""
    
    def to_dict(self):
        """Convert to dictionary for export."""
        return {
            "Date": self.date_str,
            "From": self.sender,
            "From_Email": self.sender_email,
            "To": self.recipients,
            "CC": self.cc,
            "Subject": self.subject,
            "Folder": self.folder_path,
            "Relevance_Score": self.relevance_score,
            "Matched_Companies": "; ".join(self.matched_companies),
            "Matched_Categories": "; ".join(self.matched_categories),
            "Matched_Keywords": "; ".join(self.matched_keywords[:20]),  # cap at 20 for readability
            "AI_Score": self.ai_relevance_score if self.ai_relevance_score else "",
            "AI_Summary": self.ai_summary,
            "Has_Attachments": "Yes" if self.has_attachments else "No",
            "Attachment_Names": "; ".join(self.attachment_names),
            "Body_Preview": self.body[:1000] if self.body else "",
            "Full_Body": self.body
        }


class PSTParser:
    """
    Handles opening and extracting emails from .pst files.
    
    Uses the pypff library to read Outlook PST archives without
    needing Outlook installed. Falls back to providing instructions
    for alternative methods if pypff isn't available.
    """
    
    def __init__(self, pst_path):
        self.pst_path = pst_path
        self.emails = []
        self.total_count = 0
        self.error_count = 0
    
    def parse(self):
        """Parse the PST file and extract all emails."""
        if not HAS_PYPFF:
            print("\n[ERROR] pypff library not found.")
            print("Install it with: pip install pypff")
            print("\nAlternative: Use readpst to convert your PST to mbox format:")
            print("  1. Install readpst (part of libpst package)")
            print("  2. Run: readpst -o output_folder -r your_file.pst")
            print("  3. Then use this tool with the --mbox flag instead")
            sys.exit(1)
        
        if not os.path.exists(self.pst_path):
            print(f"\n[ERROR] PST file not found: {self.pst_path}")
            sys.exit(1)
        
        file_size = os.path.getsize(self.pst_path) / (1024 * 1024)  # MB
        print(f"\nOpening PST file: {self.pst_path}")
        print(f"File size: {file_size:.1f} MB")
        print("This may take a moment for large files...\n")
        
        try:
            pst = pypff.file()
            pst.open(self.pst_path)
            root = pst.get_root_folder()
            self._process_folder(root, "")
            pst.close()
        except Exception as e:
            print(f"\n[ERROR] Failed to open PST file: {e}")
            print("\nIf the file is corrupted or encrypted, try:")
            print("  1. Opening it in Outlook and re-exporting")
            print("  2. Using scanpst.exe (Outlook's built-in repair tool)")
            print("  3. Converting to mbox with readpst")
            sys.exit(1)
        
        print(f"\nExtraction complete!")
        print(f"  Total emails found: {self.total_count}")
        print(f"  Successfully parsed: {len(self.emails)}")
        if self.error_count > 0:
            print(f"  Errors (skipped): {self.error_count}")
        
        return self.emails
    
    def _process_folder(self, folder, path):
        """Recursively process folders in the PST file."""
        try:
            folder_name = folder.name if folder.name else "Root"
            current_path = f"{path}/{folder_name}" if path else folder_name
            
            # process emails in this folder
            num_messages = folder.number_of_sub_messages
            if num_messages > 0:
                print(f"  Processing: {current_path} ({num_messages} messages)")
            
            for i in range(num_messages):
                self.total_count += 1
                try:
                    message = folder.get_sub_message(i)
                    email = self._parse_message(message, current_path)
                    if email:
                        self.emails.append(email)
                except Exception as e:
                    self.error_count += 1
                    if self.error_count <= 5:  # only show first 5 errors
                        print(f"    Warning: Could not parse message {i} in {current_path}: {e}")
                
                # progress indicator for large folders
                if self.total_count % 500 == 0:
                    print(f"    ... processed {self.total_count} emails so far")
            
            # recurse into subfolders
            for i in range(folder.number_of_sub_folders):
                try:
                    subfolder = folder.get_sub_folder(i)
                    self._process_folder(subfolder, current_path)
                except Exception as e:
                    print(f"    Warning: Could not access subfolder {i} in {current_path}: {e}")
                    
        except Exception as e:
            print(f"    Warning: Error processing folder {path}: {e}")
    
    def _parse_message(self, message, folder_path):
        """Extract data from a single email message."""
        email = EmailRecord()
        email.folder_path = folder_path
        
        # subject
        try:
            email.subject = message.subject or ""
        except:
            email.subject = ""
        
        # sender
        try:
            email.sender = message.sender_name or ""
        except:
            email.sender = ""
        
        # sender email - try multiple approaches
        try:
            # try getting the actual email address
            email.sender_email = ""
            try:
                email.sender_email = message.get_sender_email_address() or ""
            except AttributeError:
                pass
            if not email.sender_email:
                # sometimes it's embedded in the sender name as "Name <email>"
                match = re.search(r'<(.+?)>', email.sender)
                if match:
                    email.sender_email = match.group(1)
                else:
                    email.sender_email = email.sender
        except:
            email.sender_email = ""
        
        # recipients
        try:
            # pypff doesn't always expose recipients cleanly
            # we try multiple approaches
            recipients = []
            try:
                num_recipients = message.number_of_recipients if hasattr(message, 'number_of_recipients') else 0
                for r in range(num_recipients):
                    try:
                        recip = message.get_recipient(r)
                        name = recip.name if hasattr(recip, 'name') else ""
                        email_addr = recip.email_address if hasattr(recip, 'email_address') else ""
                        recipients.append(f"{name} <{email_addr}>" if email_addr else name)
                    except:
                        pass
            except:
                pass
            email.recipients = "; ".join(recipients) if recipients else ""
        except:
            email.recipients = ""
        
        # try to get To/CC from transport headers if direct access didn't work
        try:
            headers = message.transport_headers or ""
            if headers:
                if not email.recipients:
                    to_match = re.search(r'^To:\s*(.+?)(?:\r?\n\S|\r?\n\r?\n)', headers, re.MULTILINE | re.DOTALL)
                    if to_match:
                        email.recipients = to_match.group(1).strip().replace('\n', ' ').replace('\r', '')
                
                cc_match = re.search(r'^Cc:\s*(.+?)(?:\r?\n\S|\r?\n\r?\n)', headers, re.MULTILINE | re.DOTALL)
                if cc_match:
                    email.cc = cc_match.group(1).strip().replace('\n', ' ').replace('\r', '')
        except:
            pass
        
        # date
        try:
            email.date = message.delivery_time
            if email.date:
                email.date_str = email.date.strftime("%Y-%m-%d %H:%M:%S")
            else:
                email.date_str = ""
        except:
            email.date_str = ""
        
        # body - try plain text first, then HTML
        try:
            email.body = message.plain_text_body or ""
            if isinstance(email.body, bytes):
                email.body = email.body.decode('utf-8', errors='replace')
        except:
            email.body = ""
        
        if not email.body:
            try:
                html_body = message.html_body or ""
                if isinstance(html_body, bytes):
                    html_body = html_body.decode('utf-8', errors='replace')
                # strip HTML tags for searchability
                email.body = re.sub(r'<[^>]+>', ' ', html_body)
                email.body = re.sub(r'\s+', ' ', email.body).strip()
            except:
                email.body = ""
        
        # attachments
        try:
            num_attachments = message.number_of_attachments
            if num_attachments > 0:
                email.has_attachments = True
                for a in range(num_attachments):
                    try:
                        attachment = message.get_attachment(a)
                        name = attachment.name if hasattr(attachment, 'name') and attachment.name else f"attachment_{a}"
                        email.attachment_names.append(name)
                    except:
                        email.attachment_names.append(f"attachment_{a}")
        except:
            pass
        
        return email


class MboxParser:
    """
    Alternative parser for mbox format files.
    Use this if you've converted the PST with readpst first.
    
    readpst conversion command:
        readpst -o output_folder -r your_file.pst
    """
    
    def __init__(self, mbox_path):
        self.mbox_path = mbox_path
        self.emails = []
    
    def parse(self):
        """Parse mbox file(s) and extract emails."""
        import mailbox
        import email
        from email.utils import parsedate_to_datetime
        
        mbox_files = []
        path = Path(self.mbox_path)
        
        if path.is_file():
            mbox_files = [path]
        elif path.is_dir():
            # readpst creates multiple mbox files in folders
            mbox_files = list(path.rglob("*.mbox")) + list(path.rglob("*[!.]*"))
            mbox_files = [f for f in mbox_files if f.is_file()]
        
        print(f"Found {len(mbox_files)} mbox file(s) to process\n")
        
        for mbox_file in mbox_files:
            try:
                mbox = mailbox.mbox(str(mbox_file))
                folder_name = mbox_file.stem
                
                for key, message in mbox.items():
                    try:
                        record = EmailRecord()
                        record.folder_path = folder_name
                        record.subject = message.get('subject', '')
                        record.sender = message.get('from', '')
                        record.recipients = message.get('to', '')
                        record.cc = message.get('cc', '')
                        
                        # parse sender email
                        match = re.search(r'<(.+?)>', record.sender)
                        record.sender_email = match.group(1) if match else record.sender
                        
                        # parse date
                        try:
                            date_str = message.get('date', '')
                            if date_str:
                                record.date = parsedate_to_datetime(date_str)
                                record.date_str = record.date.strftime("%Y-%m-%d %H:%M:%S")
                        except:
                            record.date_str = message.get('date', '')
                        
                        # extract body
                        if message.is_multipart():
                            for part in message.walk():
                                content_type = part.get_content_type()
                                if content_type == 'text/plain':
                                    payload = part.get_payload(decode=True)
                                    if payload:
                                        record.body = payload.decode('utf-8', errors='replace')
                                    break
                            if not record.body:
                                for part in message.walk():
                                    if part.get_content_type() == 'text/html':
                                        payload = part.get_payload(decode=True)
                                        if payload:
                                            html = payload.decode('utf-8', errors='replace')
                                            record.body = re.sub(r'<[^>]+>', ' ', html)
                                            record.body = re.sub(r'\s+', ' ', record.body).strip()
                                        break
                        else:
                            payload = message.get_payload(decode=True)
                            if payload:
                                record.body = payload.decode('utf-8', errors='replace')
                        
                        self.emails.append(record)
                        
                    except Exception as e:
                        print(f"  Warning: Could not parse message in {folder_name}: {e}")
                
                print(f"  Processed: {folder_name} ({len(mbox)} messages)")
                
            except Exception as e:
                print(f"  Error reading {mbox_file}: {e}")
        
        print(f"\nTotal emails extracted: {len(self.emails)}")
        return self.emails


class EmailSearcher:
    """
    The main search engine. Takes parsed emails and finds relevant ones
    based on company list matching and keyword relevance scoring.
    """
    
    def __init__(self, emails, companies, keywords=None):
        """
        Args:
            emails: List of EmailRecord objects
            companies: List of company names/email domains to filter by
            keywords: Optional custom keyword categories dict (uses defaults if None)
        """
        self.emails = emails
        self.companies = self._prepare_companies(companies)
        self.keywords = keywords or KEYWORD_CATEGORIES
        self.results = []
    
    def _prepare_companies(self, companies):
        """
        Clean up company list for flexible matching.
        Creates both the original name and variations for searching.
        """
        prepared = []
        for company in companies:
            company = company.strip()
            if not company:
                continue
            
            entry = {
                "original": company,
                "lower": company.lower(),
                "variations": set()
            }
            
            # add the base name
            entry["variations"].add(company.lower())
            
            # if it looks like an email domain, extract the domain name
            if "@" in company:
                domain = company.split("@")[1].lower()
                entry["variations"].add(domain)
                # also add just the company part of the domain
                domain_name = domain.split(".")[0]
                entry["variations"].add(domain_name)
            
            # if it looks like a domain, extract the name part
            if "." in company and "@" not in company:
                domain_name = company.split(".")[0].lower()
                entry["variations"].add(domain_name)
            
            # strip common suffixes for matching
            for suffix in [" llc", " inc", " corp", " ltd", " co", " company",
                          " group", " partners", " services", " solutions",
                          ", llc", ", inc", ", corp", ", ltd"]:
                if company.lower().endswith(suffix):
                    stripped = company[:len(company)-len(suffix)].lower().strip()
                    entry["variations"].add(stripped)
            
            prepared.append(entry)
        
        return prepared
    
    def search(self):
        """
        Run the search across all emails.
        Returns emails sorted by relevance score (highest first).
        """
        print("\n" + "="*60)
        print("SEARCHING EMAILS")
        print("="*60)
        print(f"Emails to search: {len(self.emails)}")
        print(f"Companies to match: {len(self.companies)}")
        print(f"Keyword categories: {len(self.keywords)}")
        print()
        
        company_matches = 0
        keyword_matches = 0
        both_matches = 0
        
        for i, email in enumerate(self.emails):
            # progress indicator
            if (i + 1) % 1000 == 0:
                print(f"  Searched {i + 1}/{len(self.emails)} emails...")
            
            # step 1: check if this email involves any of the target companies
            matched_companies = self._check_companies(email)
            
            # step 2: check for relevant keywords
            score, matched_kw, matched_cats = self._score_relevance(email)
            
            if matched_companies:
                company_matches += 1
                email.matched_companies = matched_companies
            
            if score > 0:
                keyword_matches += 1
                email.matched_keywords = matched_kw
                email.matched_categories = matched_cats
                email.relevance_score = score
            
            # step 3: if it matches companies AND has relevant keywords, it's a hit
            if matched_companies and score > 0:
                both_matches += 1
                # boost score for emails that match both criteria
                email.relevance_score = score * 2
                self.results.append(email)
            elif matched_companies:
                # still include company matches but with lower priority
                email.relevance_score = 1  # minimum score
                self.results.append(email)
        
        # sort by relevance score (highest first), then by date
        self.results.sort(key=lambda e: (-e.relevance_score, e.date_str or ""), reverse=False)
        
        print(f"\n{'='*60}")
        print(f"SEARCH RESULTS SUMMARY")
        print(f"{'='*60}")
        print(f"  Emails matching target companies: {company_matches}")
        print(f"  Emails with relevant keywords: {keyword_matches}")
        print(f"  Emails matching BOTH (highest priority): {both_matches}")
        print(f"  Total results to review: {len(self.results)}")
        
        # show breakdown by category
        if self.results:
            print(f"\n  Category breakdown (among results):")
            cat_counts = defaultdict(int)
            for email in self.results:
                for cat in email.matched_categories:
                    cat_counts[cat] += 1
            for cat, count in sorted(cat_counts.items(), key=lambda x: -x[1]):
                print(f"    - {cat}: {count} emails")
        
        return self.results
    
    def _check_companies(self, email):
        """Check if an email involves any of the target companies."""
        matched = []
        
        # build a searchable text from all address fields
        address_text = " ".join([
            email.sender or "",
            email.sender_email or "",
            email.recipients or "",
            email.cc or "",
            email.bcc or ""
        ]).lower()
        
        # also check the body for company mentions
        body_lower = (email.body or "").lower()
        subject_lower = (email.subject or "").lower()
        
        for company in self.companies:
            for variation in company["variations"]:
                if variation in address_text:
                    matched.append(company["original"])
                    break
                elif variation in subject_lower or variation in body_lower:
                    matched.append(company["original"])
                    break
        
        return matched
    
    def _score_relevance(self, email):
        """
        Score an email's relevance based on keyword matches.
        Returns (score, matched_keywords, matched_categories).
        """
        score = 0
        matched_keywords = []
        matched_categories = set()
        
        # combine subject and body for searching
        searchable = f"{email.subject or ''} {email.body or ''}".lower()
        
        if not searchable.strip():
            return 0, [], []
        
        for category, config in self.keywords.items():
            weight = config["weight"]
            for term in config["terms"]:
                if term.lower() in searchable:
                    score += weight
                    matched_keywords.append(term)
                    matched_categories.add(category)
        
        return score, matched_keywords, list(matched_categories)


class AIRelevanceScorer:
    """
    Optional: Uses Claude API to perform semantic analysis of emails
    for deeper relevance scoring beyond simple keyword matching.
    
    This is the secret sauce that makes this tool competitive with
    expensive eDiscovery platforms. Instead of just matching keywords,
    it actually READS the emails and understands context.
    """
    
    def __init__(self, api_key=None):
        if not HAS_ANTHROPIC:
            print("\n[WARNING] anthropic library not installed.")
            print("Install with: pip install anthropic")
            print("AI scoring will be skipped.\n")
            self.client = None
            return
        
        api_key = api_key or os.environ.get("ANTHROPIC_API_KEY")
        if not api_key:
            print("\n[WARNING] No Anthropic API key found.")
            print("Set ANTHROPIC_API_KEY environment variable or pass --api-key")
            print("AI scoring will be skipped.\n")
            self.client = None
            return
        
        self.client = Anthropic(api_key=api_key)
        print("AI relevance scoring enabled (Claude API)\n")
    
    def score_emails(self, emails, batch_size=5):
        """
        Score a list of emails for relevance using Claude.
        Processes in batches for efficiency.
        """
        if not self.client:
            return emails
        
        print(f"\nRunning AI relevance scoring on {len(emails)} emails...")
        print("(This sends email content to Claude API for analysis)\n")
        
        scored = 0
        for i in range(0, len(emails), batch_size):
            batch = emails[i:i + batch_size]
            
            # build the prompt with the batch of emails
            email_texts = []
            for j, email in enumerate(batch):
                email_texts.append(
                    f"--- EMAIL {j+1} ---\n"
                    f"Date: {email.date_str}\n"
                    f"From: {email.sender} ({email.sender_email})\n"
                    f"To: {email.recipients}\n"
                    f"Subject: {email.subject}\n"
                    f"Body: {(email.body or '')[:2000]}\n"  # cap body length
                )
            
            prompt = f"""You are assisting a federal criminal defense attorney reviewing emails for a healthcare fraud case. The client owned a medical equipment company and hired marketing companies to sell products. The defense argument is that whenever the client learned marketing companies were not getting proper doctor authorizations, the client would terminate the relationship.

Please analyze each email below and rate its relevance to this defense on a scale of 1-10:
- 10: Directly shows client discovering authorization issues and terminating a marketing company
- 7-9: Discusses authorization problems or relationship termination with marketing companies  
- 4-6: Mentions authorizations, compliance, or marketing relationships in general
- 1-3: Tangentially related or not relevant

For each email, provide:
1. A relevance score (1-10)
2. A one-sentence summary of why it's relevant (or not)

Respond in JSON format:
[
  {{"email_number": 1, "score": 8, "summary": "Client emails marketing company about missing physician orders"}},
  ...
]

EMAILS TO ANALYZE:

{"".join(email_texts)}"""
            
            try:
                response = self.client.messages.create(
                    model="claude-sonnet-4-5-20250514",
                    max_tokens=1000,
                    messages=[{"role": "user", "content": prompt}]
                )
                
                # parse the response
                response_text = response.content[0].text
                # try to extract JSON from the response
                json_match = re.search(r'\[.*\]', response_text, re.DOTALL)
                if json_match:
                    results = json.loads(json_match.group())
                    for result in results:
                        idx = result["email_number"] - 1
                        if 0 <= idx < len(batch):
                            batch[idx].ai_relevance_score = result["score"]
                            batch[idx].ai_summary = result["summary"]
                            scored += 1
                
            except Exception as e:
                print(f"  Warning: AI scoring failed for batch starting at {i}: {e}")
            
            if (i + batch_size) % 25 == 0:
                print(f"  AI scored {min(i + batch_size, len(emails))}/{len(emails)} emails...")
        
        print(f"  AI scoring complete. Scored {scored} emails.")
        return emails


class ResultsExporter:
    """Exports search results to various formats."""
    
    def __init__(self, results, output_dir="./results"):
        self.results = results
        self.output_dir = Path(output_dir)
        self.output_dir.mkdir(parents=True, exist_ok=True)
    
    def export_excel(self, filename="email_search_results.xlsx"):
        """Export results to a formatted Excel workbook."""
        if not HAS_PANDAS:
            print("[WARNING] pandas not installed. Falling back to CSV export.")
            return self.export_csv(filename.replace('.xlsx', '.csv'))
        
        filepath = self.output_dir / filename
        
        # convert to dataframe
        data = [email.to_dict() for email in self.results]
        df = pd.DataFrame(data)
        
        # create excel writer with formatting
        with pd.ExcelWriter(filepath, engine='openpyxl') as writer:
            # main results sheet - without full body for readability
            cols_summary = [c for c in df.columns if c != "Full_Body"]
            df[cols_summary].to_excel(writer, sheet_name="Results_Summary", index=False)
            
            # high priority sheet - only high-scoring results
            high_priority = df[df["Relevance_Score"] >= 6]
            if not high_priority.empty:
                high_priority[cols_summary].to_excel(
                    writer, sheet_name="High_Priority", index=False
                )
            
            # full content sheet for detailed review
            df[["Date", "From", "To", "Subject", "Relevance_Score", 
                "Matched_Companies", "AI_Summary", "Full_Body"]].to_excel(
                writer, sheet_name="Full_Content", index=False
            )
            
            # statistics sheet
            stats = self._generate_stats(df)
            stats.to_excel(writer, sheet_name="Statistics", index=False)
        
        print(f"\n  Excel report saved: {filepath}")
        print(f"  Sheets: Results_Summary, High_Priority, Full_Content, Statistics")
        return filepath
    
    def export_csv(self, filename="email_search_results.csv"):
        """Export results to CSV format."""
        filepath = self.output_dir / filename
        
        data = [email.to_dict() for email in self.results]
        
        if not data:
            print("  No results to export.")
            return None
        
        with open(filepath, 'w', newline='', encoding='utf-8') as f:
            writer = csv.DictWriter(f, fieldnames=data[0].keys())
            writer.writeheader()
            writer.writerows(data)
        
        print(f"\n  CSV report saved: {filepath}")
        return filepath
    
    def export_summary_report(self, filename="search_report.txt"):
        """Generate a human-readable summary report."""
        filepath = self.output_dir / filename
        
        with open(filepath, 'w', encoding='utf-8') as f:
            f.write("=" * 70 + "\n")
            f.write("PST EMAIL SEARCH REPORT\n")
            f.write(f"Generated: {datetime.now().strftime('%Y-%m-%d %H:%M:%S')}\n")
            f.write("=" * 70 + "\n\n")
            
            f.write(f"Total relevant emails found: {len(self.results)}\n\n")
            
            # top results
            f.write("-" * 70 + "\n")
            f.write("TOP 25 MOST RELEVANT EMAILS\n")
            f.write("-" * 70 + "\n\n")
            
            for i, email in enumerate(self.results[:25]):
                f.write(f"#{i+1} [Score: {email.relevance_score}]\n")
                f.write(f"  Date:    {email.date_str}\n")
                f.write(f"  From:    {email.sender} ({email.sender_email})\n")
                f.write(f"  To:      {email.recipients}\n")
                f.write(f"  Subject: {email.subject}\n")
                f.write(f"  Companies: {', '.join(email.matched_companies)}\n")
                f.write(f"  Categories: {', '.join(email.matched_categories)}\n")
                if email.ai_summary:
                    f.write(f"  AI Summary: {email.ai_summary}\n")
                f.write(f"  Preview: {(email.body or '')[:300]}...\n")
                f.write("\n")
            
            # company breakdown
            f.write("-" * 70 + "\n")
            f.write("RESULTS BY COMPANY\n")
            f.write("-" * 70 + "\n\n")
            
            company_emails = defaultdict(list)
            for email in self.results:
                for company in email.matched_companies:
                    company_emails[company].append(email)
            
            for company, emails in sorted(company_emails.items(), key=lambda x: -len(x[1])):
                f.write(f"\n{company}: {len(emails)} relevant emails\n")
                for email in emails[:5]:
                    f.write(f"  - [{email.date_str}] {email.subject} (Score: {email.relevance_score})\n")
                if len(emails) > 5:
                    f.write(f"  ... and {len(emails) - 5} more\n")
        
        print(f"  Summary report saved: {filepath}")
        return filepath
    
    def _generate_stats(self, df):
        """Generate statistics about the search results."""
        stats_data = {
            "Metric": [
                "Total Results",
                "High Priority (Score >= 6)",
                "Medium Priority (Score 3-5)",
                "Low Priority (Score 1-2)",
                "Unique Companies Matched",
                "Emails with Attachments",
                "Date Range Start",
                "Date Range End",
            ],
            "Value": [
                len(df),
                len(df[df["Relevance_Score"] >= 6]),
                len(df[(df["Relevance_Score"] >= 3) & (df["Relevance_Score"] < 6)]),
                len(df[df["Relevance_Score"] < 3]),
                df["Matched_Companies"].nunique(),
                len(df[df["Has_Attachments"] == "Yes"]),
                df["Date"].min() if not df.empty else "N/A",
                df["Date"].max() if not df.empty else "N/A",
            ]
        }
        return pd.DataFrame(stats_data)


def load_companies(filepath):
    """
    Load company list from a text file.
    Expects one company name or email domain per line.
    Lines starting with # are treated as comments.
    """
    companies = []
    with open(filepath, 'r', encoding='utf-8') as f:
        for line in f:
            line = line.strip()
            if line and not line.startswith('#'):
                companies.append(line)
    
    print(f"\nLoaded {len(companies)} companies from {filepath}:")
    for c in companies:
        print(f"  - {c}")
    
    return companies


def main():
    parser = argparse.ArgumentParser(
        description="PST Email Searcher - Legal eDiscovery Tool",
        formatter_class=argparse.RawDescriptionHelpFormatter,
        epilog="""
Examples:
  # Basic search with PST file and company list
  python pst_email_searcher.py --pst client_emails.pst --companies companies.txt

  # Search with AI-powered relevance scoring
  python pst_email_searcher.py --pst client_emails.pst --companies companies.txt --ai-score

  # Use mbox format (after converting with readpst)
  python pst_email_searcher.py --mbox ./converted_emails/ --companies companies.txt

  # Specify output directory
  python pst_email_searcher.py --pst client_emails.pst --companies companies.txt --output ./case_results
        """
    )
    
    # input options
    input_group = parser.add_mutually_exclusive_group(required=True)
    input_group.add_argument('--pst', help='Path to the .pst file')
    input_group.add_argument('--mbox', help='Path to mbox file or directory (from readpst conversion)')
    
    parser.add_argument('--companies', required=True, help='Path to text file with company names (one per line)')
    parser.add_argument('--output', default='./results', help='Output directory for results (default: ./results)')
    
    # AI scoring options
    parser.add_argument('--ai-score', action='store_true', help='Enable AI-powered relevance scoring (requires Anthropic API key)')
    parser.add_argument('--api-key', help='Anthropic API key (or set ANTHROPIC_API_KEY env variable)')
    
    # output options
    parser.add_argument('--format', choices=['excel', 'csv', 'both'], default='both', help='Output format (default: both)')
    parser.add_argument('--max-results', type=int, default=None, help='Maximum number of results to export')
    
    args = parser.parse_args()
    
    print("\n" + "="*60)
    print("  PST EMAIL SEARCHER - Legal eDiscovery Tool")
    print("="*60)
    
    # load company list
    companies = load_companies(args.companies)
    
    # parse emails
    if args.pst:
        parser_obj = PSTParser(args.pst)
    else:
        parser_obj = MboxParser(args.mbox)
    
    emails = parser_obj.parse()
    
    if not emails:
        print("\n[ERROR] No emails were extracted. Check your input file.")
        sys.exit(1)
    
    # run search
    searcher = EmailSearcher(emails, companies)
    results = searcher.search()
    
    if not results:
        print("\n[INFO] No matching emails found.")
        print("Try broadening your company list or checking for typos.")
        sys.exit(0)
    
    # optional AI scoring
    if args.ai_score:
        scorer = AIRelevanceScorer(api_key=args.api_key)
        results = scorer.score_emails(results)
        # re-sort with AI scores factored in
        for email in results:
            if email.ai_relevance_score:
                email.relevance_score += email.ai_relevance_score * 2
        results.sort(key=lambda e: -e.relevance_score)
    
    # cap results if requested
    if args.max_results:
        results = results[:args.max_results]
    
    # export results
    print(f"\n{'='*60}")
    print("EXPORTING RESULTS")
    print(f"{'='*60}")
    
    exporter = ResultsExporter(results, output_dir=args.output)
    
    if args.format in ('excel', 'both'):
        exporter.export_excel()
    if args.format in ('csv', 'both'):
        exporter.export_csv()
    
    exporter.export_summary_report()
    
    print(f"\n{'='*60}")
    print("DONE!")
    print(f"{'='*60}")
    print(f"\nResults saved to: {args.output}/")
    print(f"Total relevant emails: {len(results)}")
    print("\nNext steps:")
    print("  1. Open the Excel/CSV file to review results")
    print("  2. High-priority emails are sorted to the top")
    print("  3. Check the 'High_Priority' sheet first")
    print("  4. Read the summary report for an overview")
    if not args.ai_score:
        print("\n  Tip: Run again with --ai-score for AI-powered analysis")
        print("  (requires Anthropic API key)")


if __name__ == "__main__":
    main()
