/*
 * Copyright 2014 Elpapo.
 *
 * Licensed under the Apache License, Version 2.0 (the "License");
 * you may not use this file except in compliance with the License.
 * You may obtain a copy of the License at
 *
 *      http://www.apache.org/licenses/LICENSE-2.0
 *
 * Unless required by applicable law or agreed to in writing, software
 * distributed under the License is distributed on an "AS IS" BASIS,
 * WITHOUT WARRANTIES OR CONDITIONS OF ANY KIND, either express or implied.
 * See the License for the specific language governing permissions and
 * limitations under the License.
 */

package controller;

import java.lang.ProcessBuilder.Redirect;
import javax.script.ScriptEngine;
import javax.script.ScriptEngineManager;
import javax.script.ScriptException;
import org.jruby.embed.ScriptingContainer;

/**
 *
 * @author Isaac Alcocer
 */
public class TabulaController {
    
    public void checkProcess() throws Exception {
        //*
        ProcessBuilder pb = new ProcessBuilder(
                "java",
                "-jar",
                "C:\\Users\\Elpapo\\Downloads\\jruby-complete-1.7.13.jar",
                "\"bin/jruby -v\"");
        pb.redirectOutput(Redirect.INHERIT);
        pb.redirectError(Redirect.INHERIT);
        Process p = pb.start();
        p.waitFor();
        p.destroy();
        /*/
        
        /*
        Process p = Runtime.getRuntime().exec(
              "java -Xmx500m -Xss1024k -jar C:\\Users\\Elpapo\\Downloads\\jruby-complete-1.7.13.jar "
              + "-e \"puts 'Hello'\" ");
        p.waitFor();

        String line;

        BufferedReader error = new BufferedReader(new InputStreamReader(p.getErrorStream()));
        while((line = error.readLine()) != null){
            System.out.println(line);
        }
        error.close();

        BufferedReader input = new BufferedReader(new InputStreamReader(p.getInputStream()));
        while((line=input.readLine()) != null){
            System.out.println(line);
        }

        input.close();

        OutputStream outputStream = p.getOutputStream();
        PrintStream printStream = new PrintStream(outputStream);
        printStream.println();
        printStream.flush();
        printStream.close();*/
        
    }
    
    //Jsr223 version
    private TabulaController() throws ScriptException {
        ScriptEngineManager manager = new ScriptEngineManager();
        ScriptEngine engine = manager.getEngineByName("jruby");
        engine.eval("puts 'Hello World!'");
    }
    
    //embed version
    private TabulaController(String a){
       ScriptingContainer container = new ScriptingContainer();
       container.runScriptlet(a);
    }

    public static void main(String[] args) throws ScriptException {
        new TabulaController();
    }
    
    
    
}
