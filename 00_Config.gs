// ==========================================
// BIMBOT v41.0 - 00_Config
// Centralized IDs and Global Variables
// ==========================================

const MAIN_DB_ID = "1dnXzt5w6R9YgnCGPpKoQIPO-Qf5ZrL7TqfnGmTTrj-A";      // Projects For BimBot - Alex (Projects_TL, Project_Code)
const INDEX_DB_ID = "1h8Y30nprr4Z0CvYO-CCorTZczvd-tlBYga2JTJGlh_0";     // Project number -> real Drive folder ID
const PROCESS_DB_ID = "1yENFJBgtF3oB6QK6hCG0RqJNZQzjebhQpyBlxQvD3ic";
const TEMPLATE_DB_ID = "1tNNaSHMhZr87kIz94xGLC3wwmYm4jficuLU74eCV97I";   // "BIM INFO" spreadsheet template
const DROPDOWN_DB_ID = "1zxhtXC5EN8rIxUnYaBpwPL6X03S9sUeYPI9P-IwKmik";
const USERS_DB_ID = "1VON77QKODL9-KQhCepQdu2HyqlOKvD6RgNw4hDXUteA";
const MISSIONS_DB_ID = "1KwSOMiyC6mD_zJPqZ2DD03xuywzKFaEsiW7xYZiltXM";
const KICKOFF_DOC_TEMPLATE_ID = "1P746-Z7VdjNqPeyFeUnO93KG_TXa53dgErvuBzqFl5o";
const PRESENTATION_TEMPLATE_ID = "1nn2zc27ZjFVGJLq84wUxE1okLEX-59ZAlndxSxOFg88";

// Every project resolves to its REAL Drive folder from INDEX_DB_ID.
// Set to true to block all writes (maintenance / safe testing).
const APP_READ_ONLY = false;

// Safety copy of the project spreadsheet before a save, at most once per window.
// Autosave fires every few seconds, so a copy per save would flood Drive.
const BACKUP_FOLDER_NAME = "_BIMBOT_Backups";
const BACKUP_MIN_INTERVAL_MINUTES = 60;
