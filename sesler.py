# -*- coding: utf-8 -*-
"""
Created on Tue Oct 27 15:24:39 2020

@author: musta
"""
from sys import byteorder
from array import array
from struct import pack
import pyaudio
import wave
import pandas as pd
import numpy as np
import io


THRESHOLD = 500
CHUNK_SIZE = 1024
FORMAT = pyaudio.paInt16
RATE = 44100

def is_silent(snd_data):
    return max(snd_data) < THRESHOLD

def normalize(snd_data):
    MAXIMUM = 16384
    times = float(MAXIMUM)/max(abs(i) for i in snd_data)

    r = array('h')
    for i in snd_data:
        r.append(int(i*times))
    return r

def trim(snd_data):
    def _trim(snd_data):
        snd_started = False
        r = array('h')

        for i in snd_data:
            if not snd_started and abs(i)>THRESHOLD:
                snd_started = True
                r.append(i)

            elif snd_started:
                r.append(i)
        return r

    # Trim to the left
    snd_data = _trim(snd_data)

    # Trim to the right
    snd_data.reverse()
    snd_data = _trim(snd_data)
    snd_data.reverse()
    return snd_data

def add_silence(snd_data, seconds):
    silence = [0] * int(seconds * RATE)
    r = array('h', silence)
    r.extend(snd_data)
    r.extend(silence)
    return r

def record():
    p = pyaudio.PyAudio()
    stream = p.open(format=FORMAT, channels=1, rate=RATE,
        input=True, output=True,
        frames_per_buffer=CHUNK_SIZE)

    num_silent = 0
    snd_started = False

    r = array('h')

    while 1:
        # little endian, signed short
        snd_data = array('h', stream.read(CHUNK_SIZE))
        if byteorder == 'big':
            snd_data.byteswap()
        r.extend(snd_data)

        silent = is_silent(snd_data)

        if silent and snd_started:
            num_silent += 1
        elif not silent and not snd_started:
            snd_started = True

        if snd_started and num_silent > 30:
            break

    sample_width = p.get_sample_size(FORMAT)
    stream.stop_stream()
    stream.close()
    p.terminate()

    r = normalize(r)
    r = trim(r)
    r = add_silence(r, 0.5)
    return sample_width, r


def record_to_file():
    "Records from the microphone and outputs the resulting data to 'path'"
    sample_width, data = record()
    data = pack('<' + ('h'*len(data)), *data)
    veriler=pd.read_excel('list.xlsx')
  
    satislar=veriler['Phrases'][0]
    print(satislar)
 
 
    for i in range(1,11):
        wf = wave.open("174410037_"+satislar.replace(" ","_").lower()+"_"+str(i)+'.wav', 'wb')
        wf.setnchannels(1)
        wf.setsampwidth(sample_width)
        wf.setframerate(RATE)
        wf.writeframes(data)
        wf.close()
 
   
    data = pd.read_excel("list.xlsx", index_col ="Phrases" )        
    veriler=data.drop([veriler['Phrases'][0]], inplace = True)
    veriler = data.apply(veriler, axis='columns')
  
    writer = pd.ExcelWriter('./list.xlsx', engine='xlsxwriter')
    veriler.to_excel(writer, 'Sheet1')
    writer.save()
 
    

if __name__ == '__main__':
    
    print("Konuşun")
    record_to_file()
    print("Kaydedildi")
    
    
    
    