// Importante: Per connettersi a un qualsiasi DB Access bisogna PRIMA installare la versione opportuna di AccessDatabaseEngine
// da https://www.microsoft.com/en-us/download/details.aspx?id=54920

extern crate odbc;
// Use this crate and set environmet variable RUST_LOG=odbc to see ODBC warnings
extern crate env_logger;
extern crate odbc_safe;
use odbc::*;

//#[allow(unused_imports)]        //permetto l'uso di use std altrimenti da errore
use odbc_safe::AutocommitOn;
use std::io;

// const DB: &str = "./EsempioStudenti.accdb";

fn main() {
    //inizializzo la connessione
    env_logger::init();
    //select conessione ok or error
    match connect() {
        Ok(()) => println!("Success"),
        Err(diag) => println!("Error: {}", diag),
    }
}

//funzione connessione + diagnostica
fn connect() -> std::result::Result<(), DiagnosticRecord> {
    //assegna a env = una ambiente odbc ??
    let env = create_environment_v3().map_err(|e| e.unwrap())?;


    //Importante: per leggere da excel usare:
    // "Driver={Microsoft Excel Driver (*.xls, *.xlsx, *.xlsm, *.xlsb)};DBQ=c:/my/file.xls"

    let mut buffer = String::new();
    println!("Inserire la stringa di connessione [esempio: Driver={{Microsoft Access Driver (*.mdb, *.accdb)}}; DBQ=c:/my/file.accdb]: ");
    io::stdin().read_line(&mut buffer).unwrap();
    if buffer.eq(&String::from("\r\n")) {
        buffer = "Driver={Microsoft Access Driver (*.mdb, *.accdb)}; DBQ=./EsempioStudenti.accdb;".to_owned(); // to_owned() converte &str in String
    }

    //attivo la conenessione passando il buffer
    let conn = env.connect_with_connection_string(&buffer)?;
    execute_statement(&conn)
}

//questa funzione esegue una query qualsiasi
fn execute_statement<'env>(conn: &Connection<'env, AutocommitOn>) -> Result<()> {
    // Crea un nuovo Statement (comando SQL vuoto) a partire dalla connessione passata come parametro.
    let stmt = Statement::with_parent(conn)?;

    let mut sql_text = String::new();
    println!("Inserire il comando SQL: [default: SELECT * FROM STUDENTE;]");
    io::stdin().read_line(&mut sql_text).unwrap();
    
    // Se non scrivi nulla, cioè premi invio (\r\n), viene eseguita una query di default.
    if sql_text.eq(&String::from("\r\n")) {
        sql_text = "SELECT * FROM STUDENTE;".to_owned(); // to_owned() converte &str in String
    }

    // eseguo la query scritta da linea di comando con exec_direct() 
    // costrutto match con due rami: Data(statement) e NoData(_)
    match stmt.exec_direct(&sql_text)? {
        // Se ci sono dati, li stampo in output
        Data(mut stmt) => {
            // Stampo ogni colonna separando con uno spazio
            let cols = stmt.num_result_cols()?; // numero di colonne
            // finche' c'e' qualche (Some) riga (la riga si prende con statement.fetch())
            while let Some(mut cursor) = stmt.fetch()? {
                // per ogni colonna (le colonne partono da 1 per ODBC)
                for i in 1..=cols { // intervallo di numeri 1,2,3,...,numero_colonne. l'= serve per includere l'estremo superiore dell'intervallo
                    
                    match cursor.get_data::<&str>(i as u16)? {
                        // se ci sono dati nella cella, stampo con uno spazio davanti per separare.
                        Some(val) => print!(" {}", val),
                        // se non ci sono dati nella cella, stampo la parola NULL (valore nullo)
                        None => print!(" NULL"),
                    }
                }
                // vado a capo
                println!("");
            }
        }
        // Se la query non restituisce dati, stampo una stringa.
        NoData(_) => println!("La query e' stata eseguita correttamente, ma non ha restituito dati"),
    }

    // Restituisco Ok al chiamante. Se si arriva qui, è andato tutto bene (Ok!)
    Ok(())
}
