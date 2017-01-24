/*
  jotlmsg API
  Copyright (C) 2017 Cédric Tabin

  This file is part of jotlmsg, a library to easily manipulate and generate
  Outlook msg files.
  The author can be contacted on http://www.astorm.ch/blog/index.php?contact

  Redistribution and use in source and binary forms, with or without
  modification, are permitted provided that the following conditions
  are met:
  1. Redistributions of source code must retain the above copyright
     notice, this list of conditions and the following disclaimer.
  2. Redistributions in binary form must reproduce the above copyright
     notice, this list of conditions and the following disclaimer in
     the documentation and/or other materials provided with the
     distribution.
  3. The names of the authors may not be used to endorse or promote
     products derived from this software without specific prior
     written permission.
 
  THIS SOFTWARE IS PROVIDED BY THE AUTHORS ``AS IS'' AND ANY EXPRESS
  OR IMPLIED WARRANTIES, INCLUDING, BUT NOT LIMITED TO, THE IMPLIED
  WARRANTIES OF MERCHANTABILITY AND FITNESS FOR A PARTICULAR PURPOSE
  ARE DISCLAIMED.  IN NO EVENT SHALL THE AUTHORS BE LIABLE FOR ANY
  DIRECT, INDIRECT, INCIDENTAL, SPECIAL, EXEMPLARY, OR CONSEQUENTIAL
  DAMAGES (INCLUDING, BUT NOT LIMITED TO, PROCUREMENT OF SUBSTITUTE
  GOODS OR SERVICES; LOSS OF USE, DATA, OR PROFITS; OR BUSINESS
  INTERRUPTION) HOWEVER CAUSED AND ON ANY THEORY OF LIABILITY, WHETHER
  IN CONTRACT, STRICT LIABILITY, OR TORT (INCLUDING NEGLIGENCE OR
  OTHERWISE) ARISING IN ANY WAY OUT OF THE USE OF THIS SOFTWARE, EVEN
  IF ADVISED OF THE POSSIBILITY OF SUCH DAMAGE.
*/

package ch.astorm.jotlmsg.io;

import java.io.IOException;
import java.io.OutputStream;
import java.util.List;
import org.apache.poi.hsmf.datatypes.PropertyValue;

/**
 * Handles the writing of a simple storage chunk.
 * This class is based on {@link org.apache.poi.hsmf.datatypes.StoragePropertiesChunk}.
 * 
 * @author Cedric Tabin
 */
public class StoragePropertiesChunk extends PropertiesChunk {
    
    @Override
    protected List<PropertyValue> writeHeaderData(OutputStream out) throws IOException {
        //storage header
        //page 23, point 2.4.1.3

        // 8 bytes of reserved zeros
        out.write(new byte[8]);

        return super.writeHeaderData(out);
    }
}
