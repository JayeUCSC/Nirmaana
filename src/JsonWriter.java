import org.json.simple.JSONObject;

import java.io.FileWriter;
import java.io.IOException;
import java.util.HashMap;

/**
 * Created by Amila on 2016-09-20.
 */
public class JsonWriter {
    void write(HashMap data, String path) throws IOException {

        JSONObject jsonObjectOuter = new JSONObject();
        jsonObjectOuter.putAll(data);
        //System.out.println(jsonObjectOuter);
        FileWriter fw = new FileWriter(path);
        fw.write(jsonObjectOuter.toJSONString());
        fw.flush();
        fw.close();

    }
}
