import org.apache.commons.io.FileUtils
import java.io.File
import java.io.IOException
import java.util.*
import javax.mail.*


private const val SUBJECT_TEXT = "INFO NORV"

object FetchEmails {

    @JvmStatic
    fun main(args: Array<String>) {

//        replaceImages("INFO NORVEGE - 06")
//        if (true) return

        println("Connecting to IMAP server.")
        val properties = Properties()
        properties["mail.imaps.ssl.enable"] = "true"
        val session: Session = Session.getInstance(properties)
        val store: Store = session.getStore("imaps")
        val config = Properties()
        config.load(File("config.properties").inputStream())
        store.connect("imap.gmail.com", config.getProperty("email"), config.getProperty("password"))
        val inbox: Folder = store.getFolder("INBOX")
        inbox.open(Folder.READ_ONLY)

        println("Fetching emails.")
        val messages: Array<Message> = inbox.messages

        // Write each email to a separate paragraph in the Word document
        for (message in messages) {

            // Get the email subject
            val subject = message.subject
            if (subject == null || !subject.contains(SUBJECT_TEXT)) {
                continue
            }
            val fileName = subject.replace('/', '-')

            println("\nSaving $fileName")

            // create one photo directory for each message
            File(fileName).mkdir()

            val multipart: Multipart = message.content as Multipart
            // Save text to html file
            File("$fileName.html").writeText(getText(multipart.getBodyPart(0)))
            replaceImages(fileName)

            saveAttachedPhotos(multipart, fileName)
        }

        // Close the IMAP connection
        inbox.close(false)
        store.close()
    }

    private fun saveAttachedPhotos(multipart: Multipart, fileName: String) {
        val numPhotos = multipart.count - 1
        print("Saving $numPhotos photos: ")
        for (i in 1..numPhotos) {
            val part: Part = multipart.getBodyPart(i)
            val partFileName = part.fileName.lowercase()
            if (part.isMimeType("image/jpeg")) {
                // Write the image to a file
                FileUtils.copyInputStreamToFile(part.inputStream, File("$fileName/$partFileName"))
                print(".")
                // warning if filename is not in the right format
                if (!partFileName.matches(Regex("""\d\d\.jpg"""))) {
                    println("\nWrongly named photo $partFileName")
                }
            } else {
                println("\nFound non-jpeg attachment $partFileName")
            }
        }
        println()
    }

    @Throws(MessagingException::class, IOException::class)
    private fun getText(p: Part): String {
        if (p.isMimeType("text/*")) {
            return p.content as String
        }

        if (p.isMimeType("multipart/alternative")) {
            // prefer html text over plain text
            val mp = p.content as Multipart
            val text = ""
            for (i in 0 until mp.count) {
                val bp: Part = mp.getBodyPart(i)
                if (bp.isMimeType("text/plain")) {
                    continue
                } else if (bp.isMimeType("text/html")) {
                    val s = getText(bp)
                    return s
                } else {
                    return getText(bp)
                }
            }
            return text
        } else if (p.isMimeType("multipart/*")) {
            val mp = p.content as Multipart
            for (i in 0 until mp.count) {
                val s = getText(mp.getBodyPart(i))
                return s
            }
        }

        return ""
    }

    private fun replaceImages(fileName: String) {
        println("Replacing '(photo...)' text with links in $fileName")

        val backupDirectory = "backup"
        File(backupDirectory).mkdir()  // Create backup directory if it doesn't exist

        // Extract infoNum from fileName
        val infoNum = Regex("""(\d+)$""").find(fileName)!!.groupValues[1].toInt()

        // Read HTML file content
        var fileContent = File("$fileName.html").readText(Charsets.UTF_8)

        val backupFileName = "$backupDirectory/$fileName.html"
        println("Backing up to $backupFileName")
        File("$fileName.html").copyTo(File(backupFileName), overwrite = true)

        // Replace image references
        fileContent = fileContent.replace(
            Regex("""\(photos?\s*(\s*\d+\s*)\)""", RegexOption.MULTILINE)
        ) { matchResult ->
            val photoNum = matchResult.groupValues[1].toInt()
            img(fileName, infoNum, photoNum)
        }
        fileContent = fileContent.replace(
            Regex("""\(photos?\s*(\d+)\s*et\s*(\d+)\)""", RegexOption.MULTILINE)
        ) { matchResult ->             // "(photo n et m)"
            val startNum = matchResult.groupValues[1].toInt()
            val endNum = matchResult.groupValues[2].toInt()
            img(fileName, infoNum, startNum) +img(fileName, infoNum, endNum)
        }
        fileContent = fileContent.replace(
            Regex("""\(photos?\s*(\d+)\s*à\s*(\d+)\)""", RegexOption.MULTILINE)
        ) { matchResult ->             // "(photos n et m)"
            val startNum = matchResult.groupValues[1].toInt()
            val endNum = matchResult.groupValues[2].toInt()
            val imageTags = (startNum..endNum).joinToString(separator = "") {
                img(fileName, infoNum, it)
            }
            imageTags
        }

        // Save modified content
        File("$fileName.html").writeText(fileContent, Charsets.UTF_8)
    }

    private fun img(fileName: String, infoNum: Int, photoNum: Int) =
        "<br/><br/><img src=\"$fileName/${formatNum(photoNum)}.jpg\"/><br/><br/>"

    private fun formatNum(n: Int) = n.toString().padStart(2, '0')
}
